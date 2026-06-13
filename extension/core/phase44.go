package core

import (
	"encoding/json"
	"fmt"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
)

// Phase 4.4 (PLAN.md §13): content types — rich-text cell values, in-memory
// (GD) drawings, and auto-filter column rules. Charts are PHP-side, mapped
// onto the existing native chart spec (phase3.go AddChart).

// SetRichText queues a rich-text value for a cell (JSON run list). Like other
// save-time ops it applies after the data is in the model, so the rich-text
// representation wins over any plain value streamed into the same cell.
func (w *Workbook) SetRichText(sheet, cell, jsonRuns string) error {
	if _, err := compat.TranslateRichText(jsonRuns); err != nil {
		return err
	}
	return w.queueOp(sheet, pendingOp{kind: opRichText, ref: cell, s1: jsonRuns})
}

// AddImageBytes queues an in-memory drawing (base64 PNG/JPEG/GIF). Unlike
// AddImage there is no path policy to satisfy — the bytes come from PHP (a
// GD MemoryDrawing rendered to PNG).
func (w *Workbook) AddImageBytes(sheet, cell, jsonSpec string) error {
	var spec imageSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid image spec: %w", err)
	}
	if spec.Data == "" || spec.Extension == "" {
		return fmt.Errorf("easy-excel: in-memory drawing needs data and extension")
	}
	return w.queueOp(sheet, pendingOp{kind: opImage, ref: cell, s1: jsonSpec})
}

// filterColumn is one decoded auto-filter column rule.
type filterColumn struct {
	Column     string `json:"column"`
	Expression string `json:"expression"`
}

// AutoFilterWithColumns sets an auto-filter with per-column rules. Column
// rules need the FilterColumn worksheet XML excelize emits, so this always
// goes through the model (never the streaming container patch).
func (w *Workbook) AutoFilterWithColumns(sheet, ref, columnsJSON string) error {
	var cols []filterColumn
	if err := json.Unmarshal([]byte(columnsJSON), &cols); err != nil {
		return fmt.Errorf("easy-excel: invalid auto-filter columns: %w", err)
	}
	opts := make([]excelize.AutoFilterOptions, 0, len(cols))
	for _, c := range cols {
		opts = append(opts, excelize.AutoFilterOptions{Column: c.Column, Expression: c.Expression})
	}
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	tl, br, err := splitRange(ref)
	if err != nil {
		return err
	}
	// a plain auto-filter on this sheet may have been parked for the patch;
	// the column-rule version supersedes it and must use the model
	delete(w.filters, sheet)
	if st.random() {
		if err := w.mutable(); err != nil {
			return err
		}
		return w.f.AutoFilter(sheet, tl+":"+br, opts)
	}
	st.pending = append(st.pending, pendingOp{kind: opAutoFilterCols, ref: tl + ":" + br, s1: columnsJSON})
	return nil
}

// applyOpPhase44 executes the queued wave-4.4 ops in random-access mode.
func (w *Workbook) applyOpPhase44(sheet string, op pendingOp) error {
	switch op.kind {
	case opRichText:
		runs, err := compat.TranslateRichText(op.s1)
		if err != nil {
			return err
		}
		return w.f.SetCellRichText(sheet, op.ref, runs)
	case opAutoFilterCols:
		var cols []filterColumn
		if err := json.Unmarshal([]byte(op.s1), &cols); err != nil {
			return err
		}
		opts := make([]excelize.AutoFilterOptions, 0, len(cols))
		for _, c := range cols {
			opts = append(opts, excelize.AutoFilterOptions{Column: c.Column, Expression: c.Expression})
		}
		return w.f.AutoFilter(sheet, op.ref, opts)
	}
	return fmt.Errorf("easy-excel: unknown pending op %d", op.kind)
}
