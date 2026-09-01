package core

import (
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

// Pre-computed formula cache (Phase 5 cross-cutting item 1).
//
// PhpSpreadsheet evaluates formulas and stores the result alongside the
// formula, so `<f>SUM(A1:A2)</f><v>5</v>`. excelize writes the formula only.
// Excel, LibreOffice and getCalculatedValue() all recompute on open and show
// the right answer, but a headless reader that trusts the cached <v> — which
// is what PDF/HTML rendering pipelines typically use — sees a blank cell.
//
// Two independent mitigations live here, because they address different
// readers and have very different costs:
//
//   - FullCalcOnLoad (cheap, default ON): sets the workbook's
//     calcPr/@fullCalcOnLoad flag, telling a *spreadsheet application* to
//     recalculate everything when it opens the file. Costs one XML attribute
//     and nothing at write time. Does nothing for a reader that never
//     calculates.
//   - PrecalculateFormulas (expensive, default OFF): evaluates every formula
//     cell and writes the result into the cache, so a non-calculating reader
//     shows a value instead of a blank. It is O(formula cells), forces
//     random-access mode, and so forfeits streaming — hence opt-in, and never
//     enabled by inheriting PhpSpreadsheet's default.
//
// Two limitations, both inherent to excelize and both verified rather than
// assumed (see cacheFormulaResults):
//
//   - only numeric results can be cached at all;
//   - a cached cell is emitted as `t="str"`, so its number format does not
//     apply when a non-calculating reader renders it.
//
// The second is why this is not a drop-in match for PhpSpreadsheet's
// pre-calculation, and why fullCalcOnLoad remains the default mitigation.

// SetFullCalcOnLoad toggles the workbook's fullCalcOnLoad flag.
func (w *Workbook) SetFullCalcOnLoad(on bool) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	w.fullCalcOnLoad = on
	return nil
}

// SetPrecalculateFormulas enables the save-time evaluate-and-cache pass.
func (w *Workbook) SetPrecalculateFormulas(on bool) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	w.precalculate = on
	return nil
}

// canonicalNumber normalises a raw calc result for comparison against the
// round-tripped float: excelize emits uppercase scientific notation, and a
// leading "+" is equivalent. Anything else that differs — leading zeros,
// thousands separators, trailing padding — marks the value as a formatted
// string rather than a number.
func canonicalNumber(s string) string {
	s = strings.TrimPrefix(s, "+")
	return strings.ToUpper(s)
}

// applyCalcProps writes the fullCalcOnLoad flag. Called from the save path.
//
// Writes the flag in both directions: returning early when false left a
// workbook loaded from a file that already carried fullCalcOnLoad="true"
// unable to clear it (review item 10).
func (w *Workbook) applyCalcProps() error {
	on := w.fullCalcOnLoad
	return w.f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &on})
}

// cacheFormulaResults evaluates every formula cell and stores the result
// beside the formula, so readers that do not calculate still show a value.
//
// **Numeric results only, and typed as text.** Two separate excelize limits:
//
//  1. A string or boolean result cannot be cached correctly at all:
//     SetCellStr/SetCellValue(string) store a shared-string *index* in <v>,
//     and the following SetCellFormula relabels the cell "str" while leaving
//     that index — the cell reads back as "0"/"1" instead of its text. Those
//     results, and errors (#DIV/0! and friends), are left uncached.
//
//  2. SetCellFormula ends with an unconditional `c.T, c.IS = "str", nil`
//     (excelize cell.go), and there is no public API to reset the type
//     afterwards. Writing the value *after* the formula clears the formula
//     instead. So even a cached numeric result is emitted as
//     `<c t="str"><f>…</f><v>2500</v></c>`: readable as a value, but its
//     number format will not be applied by a reader that does not recalculate
//     (a #,##0 total renders "2500", not "2,500").
//
// Point 2 is a real fidelity loss, not a cosmetic one, which is why
// pre-calculation stays opt-in and fullCalcOnLoad is the default mitigation.
//
// The caller must hold w.mu and have already flushed streams: the formulas
// have to exist in the model before they can be read back and evaluated.
func (w *Workbook) cacheFormulaResults() (int, error) {
	cached := 0
	w.cachedNumeric = map[string]map[string]bool{}
	for _, name := range w.f.GetSheetList() {
		rows, err := w.f.GetRows(name)
		if err != nil {
			return cached, err
		}
		for r := range rows {
			for c := range rows[r] {
				cell, err := excelize.CoordinatesToCellName(c+1, r+1)
				if err != nil {
					return cached, err
				}
				formula, err := w.f.GetCellFormula(name, cell)
				if err != nil || formula == "" {
					continue
				}
				// RawCellValue: without it excelize returns the *formatted*
				// value, so a #,##0 total comes back "2,500", fails ParseFloat
				// and is silently skipped — precisely the report cell this
				// cache exists for (review blocker 2).
				value, err := w.f.CalcCellValue(name, cell, excelize.Options{RawCellValue: true})
				if err != nil {
					// An unsupported function or a genuine #VALUE!: leave the
					// cell to be recomputed on open rather than baking in a
					// wrong or empty result.
					continue
				}
				// ParseFloat alone is not enough to prove the *result* is a
				// number: TEXT(A1,"0000") yields the string "0042", which
				// parses as 42 and would be cached — and rendered — as a
				// number, losing the leading zeros. excelize's calc engine
				// knows the real type internally (formulaArg.Type) but
				// CalcCellValue returns only a string, so the round trip is
				// the available check: a genuine numeric result survives
				// float→string→float unchanged, a formatted string does not.
				number, err := strconv.ParseFloat(value, 64)
				if err != nil {
					continue // non-numeric — see the docblock
				}
				canon := canonicalNumber(value)
				if strconv.FormatFloat(number, 'f', -1, 64) != canon &&
					strconv.FormatFloat(number, 'G', -1, 64) != canon {
					continue // string result that merely looks numeric
				}
				// Order matters: the value write clears the formula, so the
				// formula is restored immediately after. excelize keeps the
				// cached <v> when a formula is set on a cell that has one.
				if err := w.f.SetCellValue(name, cell, number); err != nil {
					return cached, err
				}
				if err := w.f.SetCellFormula(name, cell, formula); err != nil {
					return cached, err
				}
				if w.cachedNumeric[name] == nil {
					w.cachedNumeric[name] = map[string]bool{}
				}
				w.cachedNumeric[name][cell] = true
				cached++
			}
		}
	}
	return cached, nil
}
