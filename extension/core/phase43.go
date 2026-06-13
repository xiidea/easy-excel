package core

import (
	"encoding/json"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// Phase 4.3 (PLAN.md §13): structure editing. Row/column insertion and
// removal are random-access operations in excelize, so they degrade a
// streaming workbook exactly like reads do (the op-log replays first, which
// keeps queued style coordinates valid). Sheet views, headers/footers and
// margins ride the pending queue like other sheet-XML ops.

// shiftFilterToPending parks a container-patch auto-filter back on the
// pending queue: structural shifts would invalidate the recorded ref, the
// in-model path keeps it adjusted.
func (w *Workbook) shiftFilterToPending(sheet string, st *sheetState) {
	if ref, ok := w.filters[sheet]; ok {
		st.pending = append(st.pending, pendingOp{kind: opAutoFilter, ref: ref})
		delete(w.filters, sheet)
	}
}

// InsertRows inserts n empty rows before row (1-based).
func (w *Workbook) InsertRows(sheet string, row, n int) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if n < 1 {
		return fmt.Errorf("easy-excel: insert count must be positive")
	}
	w.shiftFilterToPending(sheet, st)
	if err := w.ensureRandom(sheet, st); err != nil {
		return err
	}
	w.closeIter(st)
	if err := w.f.InsertRows(sheet, row, n); err != nil {
		return err
	}
	if st.maxRow >= row {
		st.maxRow += n
	}
	return nil
}

// RemoveRows deletes n rows starting at row (1-based).
func (w *Workbook) RemoveRows(sheet string, row, n int) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if n < 1 {
		return fmt.Errorf("easy-excel: remove count must be positive")
	}
	w.shiftFilterToPending(sheet, st)
	if err := w.ensureRandom(sheet, st); err != nil {
		return err
	}
	w.closeIter(st)
	for i := 0; i < n; i++ {
		if err := w.f.RemoveRow(sheet, row); err != nil {
			return err
		}
	}
	if st.maxRow >= row {
		st.maxRow -= min(n, st.maxRow-row+1)
	}
	return nil
}

// InsertCols inserts n empty columns before col (1-based).
func (w *Workbook) InsertCols(sheet string, col, n int) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if n < 1 {
		return fmt.Errorf("easy-excel: insert count must be positive")
	}
	name, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return err
	}
	w.shiftFilterToPending(sheet, st)
	if err := w.ensureRandom(sheet, st); err != nil {
		return err
	}
	w.closeIter(st)
	if err := w.f.InsertCols(sheet, name, n); err != nil {
		return err
	}
	if st.maxCol >= col {
		st.maxCol += n
	}
	return nil
}

// RemoveCols deletes n columns starting at col (1-based).
func (w *Workbook) RemoveCols(sheet string, col, n int) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if n < 1 {
		return fmt.Errorf("easy-excel: remove count must be positive")
	}
	name, err := excelize.ColumnNumberToName(col)
	if err != nil {
		return err
	}
	w.shiftFilterToPending(sheet, st)
	if err := w.ensureRandom(sheet, st); err != nil {
		return err
	}
	w.closeIter(st)
	for i := 0; i < n; i++ {
		if err := w.f.RemoveCol(sheet, name); err != nil {
			return err
		}
	}
	if st.maxCol >= col {
		st.maxCol -= min(n, st.maxCol-col+1)
	}
	return nil
}

// MoveSheetTo places a sheet at a 0-based workbook position.
func (w *Workbook) MoveSheetTo(name string, index int) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	if _, err := w.state(name); err != nil {
		return err
	}
	list := w.f.GetSheetList()
	if index < 0 || index >= len(list) {
		return fmt.Errorf("easy-excel: sheet index %d out of range", index)
	}
	// list without the moving sheet, to find what currently sits at index
	others := make([]string, 0, len(list)-1)
	for _, n := range list {
		if n != name {
			others = append(others, n)
		}
	}
	if index == len(others) { // end position: move the current last before us
		if err := w.f.MoveSheet(name, others[len(others)-1]); err != nil {
			return err
		}
		return w.f.MoveSheet(others[len(others)-1], name)
	}
	return w.f.MoveSheet(name, others[index])
}

// CopySheetTo duplicates a sheet under a new name (appended; reorder with
// MoveSheetTo). The source degrades to random access first.
func (w *Workbook) CopySheetTo(from, newName string) (int, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return 0, errClosed
	}
	src, err := w.state(from)
	if err != nil {
		return 0, err
	}
	if _, exists := w.sheets[newName]; exists {
		return 0, fmt.Errorf("easy-excel: sheet %q already exists", newName)
	}
	if err := w.ensureRandom(from, src); err != nil {
		return 0, err
	}
	idx, err := w.f.NewSheet(newName)
	if err != nil {
		return 0, err
	}
	fromIdx, err := w.f.GetSheetIndex(from)
	if err != nil {
		return 0, err
	}
	if err := w.f.CopySheet(fromIdx, idx); err != nil {
		return 0, err
	}
	w.sheets[newName] = &sheetState{
		maxRow: src.maxRow, maxCol: src.maxCol, dimsKnown: src.dimsKnown,
	}
	return idx, nil
}

// sheetViewSpec mirrors the shim's SheetView / tab color payload.
type sheetViewSpec struct {
	ShowGridlines *bool    `json:"showGridlines"`
	ZoomScale     *float64 `json:"zoomScale"`
	RightToLeft   *bool    `json:"rightToLeft"`
	TabColor      string   `json:"tabColor"`
}

// SetSheetView queues gridline/zoom/RTL/tab-color settings (save-time).
func (w *Workbook) SetSheetView(sheet, jsonSpec string) error {
	var spec sheetViewSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid sheet view spec: %w", err)
	}
	return w.queueOp(sheet, pendingOp{kind: opSheetView, s1: jsonSpec})
}

// headerFooterSpec mirrors PhpSpreadsheet's HeaderFooter (placeholder codes
// like &P/&N/&D pass through unchanged).
type headerFooterSpec struct {
	OddHeader        string `json:"oddHeader"`
	OddFooter        string `json:"oddFooter"`
	EvenHeader       string `json:"evenHeader"`
	EvenFooter       string `json:"evenFooter"`
	FirstHeader      string `json:"firstHeader"`
	FirstFooter      string `json:"firstFooter"`
	DifferentFirst   bool   `json:"differentFirst"`
	DifferentOddEven bool   `json:"differentOddEven"`
}

// SetHeaderFooter queues print headers/footers (save-time).
func (w *Workbook) SetHeaderFooter(sheet, jsonSpec string) error {
	var spec headerFooterSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid header/footer spec: %w", err)
	}
	return w.queueOp(sheet, pendingOp{kind: opHeaderFooter, s1: jsonSpec})
}

// marginsSpec mirrors PhpSpreadsheet's PageMargins (inches); negative values
// mean "leave unchanged".
type marginsSpec struct {
	Top    float64 `json:"top"`
	Bottom float64 `json:"bottom"`
	Left   float64 `json:"left"`
	Right  float64 `json:"right"`
	Header float64 `json:"header"`
	Footer float64 `json:"footer"`
}

// SetPageMargins queues print margins (save-time).
func (w *Workbook) SetPageMargins(sheet, jsonSpec string) error {
	var spec marginsSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid margins spec: %w", err)
	}
	return w.queueOp(sheet, pendingOp{kind: opMargins, s1: jsonSpec})
}

// applyOpPhase43 executes the queued wave-4.3 ops in random-access mode.
func (w *Workbook) applyOpPhase43(sheet string, op pendingOp) error {
	switch op.kind {
	case opSheetView:
		var spec sheetViewSpec
		if err := json.Unmarshal([]byte(op.s1), &spec); err != nil {
			return err
		}
		opts := excelize.ViewOptions{
			ShowGridLines: spec.ShowGridlines,
			ZoomScale:     spec.ZoomScale,
			RightToLeft:   spec.RightToLeft,
		}
		if err := w.f.SetSheetView(sheet, 0, &opts); err != nil {
			return err
		}
		if spec.TabColor != "" {
			color := spec.TabColor
			return w.f.SetSheetProps(sheet, &excelize.SheetPropsOptions{TabColorRGB: &color})
		}
		return nil
	case opHeaderFooter:
		var spec headerFooterSpec
		if err := json.Unmarshal([]byte(op.s1), &spec); err != nil {
			return err
		}
		return w.f.SetHeaderFooter(sheet, &excelize.HeaderFooterOptions{
			OddHeader:        spec.OddHeader,
			OddFooter:        spec.OddFooter,
			EvenHeader:       spec.EvenHeader,
			EvenFooter:       spec.EvenFooter,
			FirstHeader:      spec.FirstHeader,
			FirstFooter:      spec.FirstFooter,
			DifferentFirst:   spec.DifferentFirst,
			DifferentOddEven: spec.DifferentOddEven,
		})
	case opMargins:
		var spec marginsSpec
		if err := json.Unmarshal([]byte(op.s1), &spec); err != nil {
			return err
		}
		opts := excelize.PageLayoutMarginsOptions{}
		set := func(dst **float64, v float64) {
			if v >= 0 {
				val := v
				*dst = &val
			}
		}
		set(&opts.Top, spec.Top)
		set(&opts.Bottom, spec.Bottom)
		set(&opts.Left, spec.Left)
		set(&opts.Right, spec.Right)
		set(&opts.Header, spec.Header)
		set(&opts.Footer, spec.Footer)
		return w.f.SetPageMargins(sheet, &opts)
	}
	return fmt.Errorf("easy-excel: unknown pending op %d", op.kind)
}
