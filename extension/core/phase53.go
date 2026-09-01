package core

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// Wave 5.3: manual page breaks and cell selection.
//
// Both are save-time ops, like the rest of the non-streamable surface
// (COMPAT.md divergence 11): excelize writes them into the worksheet model,
// which the StreamWriter does not carry, so they queue and apply on flush.

// Break kinds, mirroring PhpSpreadsheet's Worksheet::BREAK_* constants.
const (
	breakNone   = 0
	breakRow    = 1
	breakColumn = 2
)

// SetBreak queues a manual page break at cell. breakType uses the
// PhpSpreadsheet BREAK_* values; BREAK_NONE removes a break at that cell.
//
// excelize models row and column breaks through one call each way, deriving
// the axis from the cell reference, so the distinction between BREAK_ROW and
// BREAK_COLUMN is carried here and resolved at apply time.
func (w *Workbook) SetBreak(sheet, cell string, breakType int) error {
	switch breakType {
	case breakNone, breakRow, breakColumn:
	default:
		return fmt.Errorf("easy-excel: unsupported break type %d", breakType)
	}
	if _, _, err := excelize.CellNameToCoordinates(cell); err != nil {
		return fmt.Errorf("easy-excel: invalid break cell %q: %w", cell, err)
	}
	return w.queueOp(sheet, pendingOp{kind: opPageBreak, ref: cell, a: breakType})
}

// SetSelection queues the selected range and active cell for a sheet,
// matching Worksheet::setSelectedCells().
//
// excelize carries selection inside the pane model rather than ViewOptions, so
// this merges into whatever pane state the sheet already has (a freeze set by
// freezePane, or an empty pane record) instead of replacing it — otherwise
// selecting a cell would silently unfreeze the header rows.
func (w *Workbook) SetSelection(sheet, ref string) error {
	if ref == "" {
		return nil
	}
	tl := ref
	if r, _, err := splitRange(ref); err == nil {
		tl = r
	}
	if _, _, err := excelize.CellNameToCoordinates(tl); err != nil {
		return fmt.Errorf("easy-excel: invalid selection %q: %w", ref, err)
	}
	return w.queueOp(sheet, pendingOp{kind: opSelection, ref: ref, s1: tl})
}

// applyOpPhase53 executes the queued wave-5.3 ops in random-access mode.
func (w *Workbook) applyOpPhase53(sheet string, op pendingOp) error {
	switch op.kind {
	case opPageBreak:
		if op.a == breakNone {
			return w.f.RemovePageBreak(sheet, op.ref)
		}
		// excelize derives the axis from the reference: a break at "A24"
		// splits above row 24, one at "O1" splits left of column O. Normalise
		// so the caller's intent survives whichever cell they passed.
		cell, err := breakCell(op.ref, op.a)
		if err != nil {
			return err
		}
		return w.f.InsertPageBreak(sheet, cell)
	case opSelection:
		panes, err := w.panesWithSelection(sheet, op.ref, op.s1)
		if err != nil {
			return err
		}
		return w.f.SetPanes(sheet, panes)
	}
	return fmt.Errorf("easy-excel: unhandled op kind %d", op.kind)
}

// breakCell normalises a break reference to the axis the caller asked for: a
// row break anchors at column A of that row, a column break at row 1 of that
// column. Without this, setBreak("O24", BREAK_ROW) would also carry a column
// split, which is not what PhpSpreadsheet means.
func breakCell(ref string, breakType int) (string, error) {
	col, row, err := excelize.CellNameToCoordinates(ref)
	if err != nil {
		return "", err
	}
	if breakType == breakRow {
		return excelize.CoordinatesToCellName(1, row)
	}
	return excelize.CoordinatesToCellName(col, 1)
}

// panesWithSelection folds a selection into the sheet's existing pane state so
// freeze panes and selection coexist.
//
// The pane state is read back from the file rather than from st.prePanes:
// queued panes are flushed (and prePanes cleared) before pending ops run, so
// by the time this executes the freeze lives only in the worksheet model. A
// naive SetPanes here would drop it and silently unfreeze the header rows.
func (w *Workbook) panesWithSelection(sheet, ref, activeCell string) (*excelize.Panes, error) {
	sel := excelize.Selection{SQRef: ref, ActiveCell: activeCell, Pane: "topLeft"}

	panes := &excelize.Panes{}
	if existing, err := w.f.GetPanes(sheet); err == nil {
		panes.Freeze = existing.Freeze
		// excelize's GetPanes never populates Split — it derives Freeze from
		// state == "frozen" and leaves Split zero — so copying it would
		// silently convert a split-pane sheet to no-split. Infer it: a pane
		// record with a split offset that is not frozen is a split pane.
		panes.Split = !existing.Freeze && (existing.XSplit > 0 || existing.YSplit > 0)
		panes.XSplit = existing.XSplit
		panes.YSplit = existing.YSplit
		panes.TopLeftCell = existing.TopLeftCell
		panes.ActivePane = existing.ActivePane
		if existing.ActivePane != "" {
			sel.Pane = existing.ActivePane
		}
	}
	panes.Selection = []excelize.Selection{sel}

	return panes, nil
}
