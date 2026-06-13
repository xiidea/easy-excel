package core

import (
	"encoding/json"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
)

// Phase 4.2 (PLAN.md §13): style read-back, validation/conditional/defined-
// name introspection, and the workbook default style. Iterators and read
// filters are PHP-side.

// GetStyleSpec returns a cell's effective style as a PhpSpreadsheet spec
// (JSON), with the workbook default layered underneath. Streaming sheets are
// answered from the style log — read-back must NOT degrade a workbook that
// is mid-write.
func (w *Workbook) GetStyleSpec(sheet, cell string) (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return "", errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return "", err
	}
	var spec compat.StyleSpec
	if !st.random() {
		col, row, err := excelize.CellNameToCoordinates(cell)
		if err != nil {
			return "", err
		}
		spec = w.foldBase()
		for i := range st.styleLog {
			if e := &st.styleLog[i]; e.containsCell(row, col) {
				spec = compat.MergeSpec(spec, e.spec)
			}
		}
	} else {
		id, err := w.f.GetCellStyle(sheet, cell)
		if err != nil {
			return "", err
		}
		style, err := w.f.GetStyle(id)
		if err != nil {
			return "", err
		}
		spec = compat.ReverseStyle(style)
		if w.defaultSpec != nil {
			spec = compat.MergeSpec(w.defaultSpec, spec)
		}
	}
	encoded, err := json.Marshal(spec)
	if err != nil {
		return "", err
	}
	return string(encoded), nil
}

// ValidationsJSON returns the sheet's data validations as a JSON list of
// {sqref, spec} pairs.
func (w *Workbook) ValidationsJSON(sheet string) (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return "", errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return "", err
	}
	out := []map[string]any{}
	if !st.random() {
		// answered from the pending queue: no degrade mid-write
		for _, op := range st.pending {
			if op.kind != opValidation {
				continue
			}
			var spec map[string]any
			if err := json.Unmarshal([]byte(op.s1), &spec); err == nil {
				out = append(out, map[string]any{"sqref": op.ref, "spec": spec})
			}
		}
	} else {
		dvs, err := w.f.GetDataValidations(sheet)
		if err != nil {
			return "", err
		}
		for _, dv := range dvs {
			ref, spec := compat.ReverseValidation(dv)
			out = append(out, map[string]any{"sqref": ref, "spec": spec})
		}
	}
	encoded, err := json.Marshal(out)
	if err != nil {
		return "", err
	}
	return string(encoded), nil
}

// ConditionalsJSON returns the sheet's conditional formats as JSON keyed by
// range, each value the shim's rule list shape.
func (w *Workbook) ConditionalsJSON(sheet string) (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return "", errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return "", err
	}
	out := map[string][]map[string]any{}
	if !st.random() {
		// the pending payload is already the shim's rule shape
		for _, op := range st.pending {
			if op.kind != opConditional {
				continue
			}
			var rules []map[string]any
			if err := json.Unmarshal([]byte(op.s1), &rules); err == nil {
				out[op.ref] = rules
			}
		}
	} else {
		formats, err := w.f.GetConditionalFormats(sheet)
		if err != nil {
			return "", err
		}
		for ref, opts := range formats {
			rules := make([]map[string]any, 0, len(opts))
			for _, opt := range opts {
				var style *excelize.Style
				if opt.Format != nil {
					if s, err := w.f.GetConditionalStyle(*opt.Format); err == nil {
						style = s
					}
				}
				rules = append(rules, compat.ReverseConditional(opt, style))
			}
			out[ref] = rules
		}
	}
	encoded, err := json.Marshal(out)
	if err != nil {
		return "", err
	}
	return string(encoded), nil
}

// DefinedNamesJSON lists defined names as [{name, refersTo, scope}].
func (w *Workbook) DefinedNamesJSON() (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return "", errClosed
	}
	out := []map[string]string{}
	for _, dn := range w.f.GetDefinedName() {
		out = append(out, map[string]string{
			"name": dn.Name, "refersTo": dn.RefersTo, "scope": dn.Scope,
		})
	}
	encoded, err := json.Marshal(out)
	if err != nil {
		return "", err
	}
	return string(encoded), nil
}

// SetDefaultStyle sets the workbook default style (getDefaultStyle()). It is
// layered under every style fold and materialized as a full-width column
// style so untouched cells render with it: through the StreamWriter for
// sheets that have not started streaming, directly for random-access sheets.
func (w *Workbook) SetDefaultStyle(jsonSpec string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	spec, err := compat.ParseStyleSpec(jsonSpec)
	if err != nil {
		return err
	}
	if _, err := compat.TranslateStyle(spec); err != nil {
		return err
	}
	w.defaultSpec = spec
	for name, st := range w.sheets {
		if st.random() {
			if err := w.mutable(); err != nil {
				return err
			}
			if err := w.applyDefaultColStyle(name); err != nil {
				return err
			}
			continue
		}
		st.preDefault = true
	}
	return nil
}

func (w *Workbook) applyDefaultColStyle(sheet string) error {
	id, err := w.styles.specID(w.f, w.defaultSpec)
	if err != nil {
		return err
	}
	return w.f.SetColStyle(sheet, "A:XFD", id)
}

// foldBase is the spec folds start from: the workbook default when set.
func (w *Workbook) foldBase() compat.StyleSpec {
	if w.defaultSpec == nil {
		return compat.StyleSpec{}
	}
	return compat.MergeSpec(compat.StyleSpec{}, w.defaultSpec)
}
