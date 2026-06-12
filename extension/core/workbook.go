// Package core adapts excelize to PhpSpreadsheet semantics behind the flat
// handle ABI (PLAN.md §4). The central design is the dual write path:
//
//   - streaming mode (default for fresh workbooks): ascending row writes go
//     through excelize's StreamWriter at constant memory;
//   - random-access mode: excelize's in-memory model, used for loaded files
//     and after the one-time documented degrade triggered by out-of-order
//     writes, reads, or styling of an already-streamed sheet.
//
// The degrade serializes the workbook to memory and reopens it, so it is
// correct but costs O(file size); the PHP shim's write-behind buffering makes
// it rare in practice.
package core

import (
	"bytes"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"
	"sync"

	"github.com/xuri/excelize/v2"

	"github.com/ronisaha/easy-excel/extension/compat"
	"github.com/ronisaha/easy-excel/extension/exio"
	"github.com/ronisaha/easy-excel/extension/limits"
)

// GetMode selects what GetCell returns, mirroring PhpSpreadsheet's
// getValue / getFormattedValue / getCalculatedValue.
type GetMode byte

const (
	GetRaw GetMode = iota
	GetFormatted
	GetCalculated
)

// Safe unzip defaults (PLAN.md §8): 1GiB total, 16MiB per XML part.
const (
	unzipSizeLimit    = 1 << 30
	unzipXMLSizeLimit = 16 << 20
	baseEstimate      = 8 << 20 // accounting floor per live workbook
)

var errClosed = errors.New("easy-excel: workbook is closed")

type sheetState struct {
	sw        *excelize.StreamWriter // non-nil while streaming
	eligible  bool                   // may still enter/continue streaming
	lastRow   int                    // highest row written through sw
	maxRow    int                    // tracked dimensions (writes + opened file)
	maxCol    int
	dimsKnown bool           // false until tracked dims are trustworthy (lazy scan)
	iter      *excelize.Rows // cached forward read iterator
	iterNext  int            // 1-based row the iterator yields next
	iterRaw   bool

	// Phase-2 structure logs (see structure.go for the lifecycle)
	styleLog   []styleEntry
	rowHeights map[int]*heightEntry
	preWidths  []colWidthOp
	prePanes   *excelize.Panes
	preMerges  [][2]string
	pending    []pendingOp
}

// Workbook wraps one excelize file plus per-sheet streaming state.
// All methods are safe for concurrent use; a single mutex confines the
// workbook to one operation at a time (PLAN.md §7.4).
type Workbook struct {
	mu       sync.Mutex
	f        *excelize.File
	sheets   map[string]*sheetState
	gate     *limits.Gate
	policy   *exio.Policy
	styles   styleInterner
	estBytes int64
	degraded bool
	closed   bool
}

// Env wires the process-wide gate and path policy into workbooks.
type Env struct {
	Gate   *limits.Gate
	Policy *exio.Policy
}

func (e *Env) gate() *limits.Gate {
	if e == nil || e.Gate == nil {
		panic("easy-excel: core.Env requires a limits.Gate")
	}
	return e.Gate
}

func (e *Env) policy() *exio.Policy {
	if e == nil || e.Policy == nil {
		p, _ := exio.NewPolicy(nil)
		return p
	}
	return e.Policy
}

// New creates an empty workbook with one sheet named "Worksheet",
// matching `new PhpOffice\PhpSpreadsheet\Spreadsheet()`.
func New(env *Env) (*Workbook, error) {
	if err := env.gate().ReserveMemory(baseEstimate); err != nil {
		return nil, err
	}
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "Worksheet"); err != nil {
		env.gate().ReserveMemory(-baseEstimate)
		return nil, err
	}
	w := &Workbook{
		f:        f,
		sheets:   map[string]*sheetState{"Worksheet": {eligible: true, dimsKnown: true}},
		gate:     env.gate(),
		policy:   env.policy(),
		estBytes: baseEstimate,
	}
	return w, nil
}

// Open loads an existing workbook; all sheets start in random-access mode.
func Open(path string, env *Env) (*Workbook, error) {
	abs, err := env.policy().Resolve(path)
	if err != nil {
		return nil, err
	}
	release, err := env.gate().AcquireHeavy()
	if err != nil {
		return nil, err
	}
	defer release()

	est := int64(baseEstimate)
	if fi, err := os.Stat(abs); err == nil && fi.Size()*4 > est {
		est = fi.Size() * 4
	}
	if err := env.gate().ReserveMemory(est); err != nil {
		return nil, err
	}
	f, err := excelize.OpenFile(abs, excelize.Options{
		UnzipSizeLimit:    unzipSizeLimit,
		UnzipXMLSizeLimit: unzipXMLSizeLimit,
	})
	if err != nil {
		env.gate().ReserveMemory(-est)
		return nil, fmt.Errorf("easy-excel: open %s: %w", path, err)
	}
	w := &Workbook{
		f:        f,
		sheets:   map[string]*sheetState{},
		gate:     env.gate(),
		policy:   env.policy(),
		estBytes: est,
		degraded: true, // random-access from the start
	}
	for _, name := range f.GetSheetList() {
		st := &sheetState{}
		st.maxRow, st.maxCol = dimensionOf(f, name)
		// excelize-written files often carry a degenerate <dimension
		// ref="A1"/>; only trust non-trivial dimensions, otherwise scan
		// lazily on first Dimensions() call
		st.dimsKnown = st.maxRow > 1 || st.maxCol > 1
		if !st.dimsKnown {
			st.maxRow, st.maxCol = 0, 0
		}
		w.sheets[name] = st
	}
	return w, nil
}

func dimensionOf(f *excelize.File, sheet string) (maxRow, maxCol int) {
	dim, err := f.GetSheetDimension(sheet)
	if err != nil || dim == "" {
		return 0, 0
	}
	parts := strings.Split(dim, ":")
	corner := parts[len(parts)-1]
	col, row, err := excelize.CellNameToCoordinates(corner)
	if err != nil {
		return 0, 0
	}
	return row, col
}

func (w *Workbook) state(sheet string) (*sheetState, error) {
	st, ok := w.sheets[sheet]
	if !ok {
		return nil, fmt.Errorf("easy-excel: sheet %q does not exist", sheet)
	}
	return st, nil
}

// --- sheet management -----------------------------------------------------

// AddSheet appends a sheet and returns its 0-based index.
func (w *Workbook) AddSheet(name string) (int, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return 0, errClosed
	}
	if _, exists := w.sheets[name]; exists {
		return 0, fmt.Errorf("easy-excel: sheet %q already exists", name)
	}
	idx, err := w.f.NewSheet(name)
	if err != nil {
		return 0, err
	}
	w.sheets[name] = &sheetState{eligible: !w.degraded, dimsKnown: true}
	return idx, nil
}

// DeleteSheet removes a sheet (the last sheet cannot be removed).
func (w *Workbook) DeleteSheet(name string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(name)
	if err != nil {
		return err
	}
	if len(w.sheets) == 1 {
		return errors.New("easy-excel: cannot remove the only sheet")
	}
	w.closeIter(st)
	if st.sw != nil {
		_ = st.sw.Flush()
	}
	if err := w.f.DeleteSheet(name); err != nil {
		return err
	}
	delete(w.sheets, name)
	return nil
}

// RenameSheet renames a sheet, keeping its streaming state.
func (w *Workbook) RenameSheet(oldName, newName string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(oldName)
	if err != nil {
		return err
	}
	if oldName == newName {
		return nil
	}
	if _, exists := w.sheets[newName]; exists {
		return fmt.Errorf("easy-excel: sheet %q already exists", newName)
	}
	if err := w.f.SetSheetName(oldName, newName); err != nil {
		return err
	}
	delete(w.sheets, oldName)
	w.sheets[newName] = st
	return nil
}

// Sheets returns sheet names in workbook order.
func (w *Workbook) Sheets() []string {
	w.mu.Lock()
	defer w.mu.Unlock()
	return w.f.GetSheetList()
}

// SetActiveSheet selects the active sheet by 0-based position.
func (w *Workbook) SetActiveSheet(index int) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	list := w.f.GetSheetList()
	if index < 0 || index >= len(list) {
		return fmt.Errorf("easy-excel: sheet index %d out of range", index)
	}
	idx, err := w.f.GetSheetIndex(list[index])
	if err != nil {
		return err
	}
	w.f.SetActiveSheet(idx)
	return nil
}

// ActiveSheet returns the 0-based position and name of the active sheet.
func (w *Workbook) ActiveSheet() (int, string) {
	w.mu.Lock()
	defer w.mu.Unlock()
	list := w.f.GetSheetList()
	active := w.f.GetActiveSheetIndex()
	for pos, name := range list {
		if idx, err := w.f.GetSheetIndex(name); err == nil && idx == active {
			return pos, name
		}
	}
	if len(list) > 0 {
		return 0, list[0]
	}
	return 0, ""
}

// --- write path -------------------------------------------------------------

// WriteRows writes a batch of rows anchored at (startCol, startRow), the hot
// path called by the PHP shim's write-behind buffer. Cells with Kind == Skip
// are left untouched in random mode and written empty in streaming mode.
func (w *Workbook) WriteRows(sheet string, startRow, startCol int, rows [][]compat.Cell) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	if startRow < 1 || startCol < 1 {
		return fmt.Errorf("easy-excel: invalid anchor row %d col %d", startRow, startCol)
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if st.eligible && startRow > st.lastRow {
		if err := w.streamRows(sheet, st, startRow, startCol, rows); err != nil {
			return err
		}
	} else {
		if err := w.ensureRandom(sheet, st); err != nil {
			return err
		}
		if err := w.randomRows(sheet, startRow, startCol, rows); err != nil {
			return err
		}
	}
	endRow := startRow + len(rows) - 1
	if endRow > st.maxRow {
		st.maxRow = endRow
	}
	for _, r := range rows {
		if c := startCol + len(r) - 1; c > st.maxCol {
			st.maxCol = c
		}
	}
	return nil
}

func (w *Workbook) streamRows(sheet string, st *sheetState, startRow, startCol int, rows [][]compat.Cell) error {
	if st.sw == nil {
		sw, err := w.f.NewStreamWriter(sheet)
		if err != nil {
			return err
		}
		if err := applyPreOps(sw, st); err != nil {
			return err
		}
		st.sw = sw
	}
	styler := w.newInlineStyler(st)
	for i, row := range rows {
		rowNum := startRow + i
		if rowNum <= st.lastRow {
			return fmt.Errorf("easy-excel: stream rows must ascend (row %d after %d)", rowNum, st.lastRow)
		}
		values := make([]interface{}, len(row))
		for j, c := range row {
			styleID := 0
			if styler != nil {
				id, err := styler.styleID(rowNum, startCol+j)
				if err != nil {
					return err
				}
				styleID = id
			}
			switch c.Kind {
			case compat.Skip:
				if styleID != 0 {
					values[j] = excelize.Cell{StyleID: styleID}
				} else {
					values[j] = nil
				}
			case compat.Str:
				if styleID != 0 {
					values[j] = excelize.Cell{StyleID: styleID, Value: c.Str}
				} else {
					values[j] = c.Str
				}
			case compat.Num:
				if styleID != 0 {
					values[j] = excelize.Cell{StyleID: styleID, Value: c.Num}
				} else {
					values[j] = c.Num
				}
			case compat.Boolean:
				if styleID != 0 {
					values[j] = excelize.Cell{StyleID: styleID, Value: c.Bool}
				} else {
					values[j] = c.Bool
				}
			case compat.Formula:
				values[j] = excelize.Cell{StyleID: styleID, Formula: c.Str}
			}
		}
		anchor, err := excelize.CoordinatesToCellName(startCol, rowNum)
		if err != nil {
			return err
		}
		var opts []excelize.RowOpts
		if h, ok := st.rowHeights[rowNum]; ok && !h.done {
			opts = append(opts, excelize.RowOpts{Height: h.height})
			h.done = true
		}
		if err := st.sw.SetRow(anchor, values, opts...); err != nil {
			return err
		}
		st.lastRow = rowNum
	}
	st.markStreamed()
	return nil
}

func (w *Workbook) randomRows(sheet string, startRow, startCol int, rows [][]compat.Cell) error {
	for i, row := range rows {
		for j, c := range row {
			if c.Kind == compat.Skip {
				continue
			}
			axis, err := excelize.CoordinatesToCellName(startCol+j, startRow+i)
			if err != nil {
				return err
			}
			if err := w.setCellLocked(sheet, axis, c); err != nil {
				return err
			}
		}
	}
	return nil
}

// SetCell writes one cell (the explicit/slow path).
func (w *Workbook) SetCell(sheet, axis string, c compat.Cell) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if err := w.ensureRandom(sheet, st); err != nil {
		return err
	}
	if c.Kind == compat.Skip {
		return nil
	}
	if err := w.setCellLocked(sheet, axis, c); err != nil {
		return err
	}
	if col, row, err := excelize.CellNameToCoordinates(axis); err == nil {
		if row > st.maxRow {
			st.maxRow = row
		}
		if col > st.maxCol {
			st.maxCol = col
		}
	}
	return nil
}

func (w *Workbook) setCellLocked(sheet, axis string, c compat.Cell) error {
	switch c.Kind {
	case compat.Str:
		return w.f.SetCellStr(sheet, axis, c.Str)
	case compat.Num:
		return w.f.SetCellValue(sheet, axis, c.Num)
	case compat.Boolean:
		return w.f.SetCellBool(sheet, axis, c.Bool)
	case compat.Formula:
		return w.f.SetCellFormula(sheet, axis, c.Str)
	}
	return nil
}

// --- degrade ---------------------------------------------------------------

// ensureRandom leaves streaming mode for the whole workbook. If any rows were
// streamed the workbook is serialized to memory and reopened (the one-time
// documented de-optimization, PLAN.md §4.2); otherwise it is a free flag flip.
func (w *Workbook) ensureRandom(sheet string, st *sheetState) error {
	if !st.eligible && st.sw == nil {
		return nil
	}
	anyStreamed := false
	for _, s := range w.sheets {
		if s.sw != nil {
			anyStreamed = true
		}
	}
	if !anyStreamed {
		for _, s := range w.sheets {
			s.eligible = false
		}
		return w.replayAll()
	}
	if err := w.gate.ReserveMemory(w.estBytes); err != nil {
		return err
	}
	w.estBytes *= 2
	for _, s := range w.sheets {
		if s.sw != nil {
			if err := s.sw.Flush(); err != nil {
				return err
			}
			s.sw = nil
		}
		s.eligible = false
		w.closeIter(s)
	}
	var buf bytes.Buffer
	if err := w.f.Write(&buf); err != nil {
		return err
	}
	if err := w.f.Close(); err != nil {
		return err
	}
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()), excelize.Options{
		UnzipSizeLimit:    unzipSizeLimit,
		UnzipXMLSizeLimit: unzipXMLSizeLimit,
	})
	if err != nil {
		return fmt.Errorf("easy-excel: degrade reopen: %w", err)
	}
	w.f = f
	w.styles.reset()
	w.degraded = true
	return w.replayAll()
}

// ensureRandomAll leaves streaming mode for every sheet and replays queued
// structure work; used by save when pending ops cannot be streamed.
func (w *Workbook) ensureRandomAll() error {
	for name, st := range w.sheets {
		if !st.random() {
			return w.ensureRandom(name, st)
		}
	}
	return w.replayAll()
}

// Degraded reports whether the workbook left streaming mode (test/metrics hook).
func (w *Workbook) Degraded() bool {
	w.mu.Lock()
	defer w.mu.Unlock()
	return w.degraded
}

// --- read path ---------------------------------------------------------------

// GetCell reads one cell. Raw mode returns the formula source (with leading
// '=') for formula cells, native types for numbers and booleans, and strings
// otherwise — matching Cell::getValue().
func (w *Workbook) GetCell(sheet, axis string, mode GetMode) (any, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return nil, errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return nil, err
	}
	if err := w.ensureRandom(sheet, st); err != nil {
		return nil, err
	}
	switch mode {
	case GetFormatted:
		return w.f.GetCellValue(sheet, axis)
	case GetCalculated:
		return w.f.CalcCellValue(sheet, axis)
	}
	if formula, err := w.f.GetCellFormula(sheet, axis); err == nil && formula != "" {
		return "=" + formula, nil
	}
	raw, err := w.f.GetCellValue(sheet, axis, excelize.Options{RawCellValue: true})
	if err != nil {
		return nil, err
	}
	return typedValue(w.f, sheet, axis, raw), nil
}

func typedValue(f *excelize.File, sheet, axis, raw string) any {
	ct, err := f.GetCellType(sheet, axis)
	if err != nil {
		return raw
	}
	switch ct {
	case excelize.CellTypeBool:
		return raw == "1" || strings.EqualFold(raw, "true")
	case excelize.CellTypeNumber, excelize.CellTypeUnset:
		if raw == "" {
			return nil
		}
		var n float64
		if _, err := fmt.Sscanf(raw, "%g", &n); err == nil {
			return n
		}
		return raw
	default:
		return raw
	}
}

// ReadRows returns up to maxRows rows starting at startRow (1-based) as
// formatted or raw strings. Sequential chunked reads reuse a cached forward
// iterator, so a full-sheet scan in chunks stays O(n) (PLAN.md §6).
// The second return value is false when the sheet is exhausted.
func (w *Workbook) ReadRows(sheet string, startRow, maxRows int, raw bool) ([][]string, bool, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return nil, false, errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return nil, false, err
	}
	if err := w.ensureRandom(sheet, st); err != nil {
		return nil, false, err
	}
	if st.iter == nil || startRow < st.iterNext || st.iterRaw != raw {
		w.closeIter(st)
		iter, err := w.f.Rows(sheet)
		if err != nil {
			return nil, false, err
		}
		st.iter, st.iterNext, st.iterRaw = iter, 1, raw
	}
	for st.iterNext < startRow {
		if !st.iter.Next() {
			w.closeIter(st)
			return nil, false, nil
		}
		st.iterNext++
	}
	var opts []excelize.Options
	if raw {
		opts = append(opts, excelize.Options{RawCellValue: true})
	}
	out := make([][]string, 0, maxRows)
	more := true
	for len(out) < maxRows {
		if !st.iter.Next() {
			more = false
			w.closeIter(st)
			break
		}
		cols, err := st.iter.Columns(opts...)
		if err != nil {
			return nil, false, err
		}
		st.iterNext++
		out = append(out, cols)
	}
	return out, more, nil
}

func (w *Workbook) closeIter(st *sheetState) {
	if st.iter != nil {
		_ = st.iter.Close()
		st.iter = nil
		st.iterNext = 0
	}
}

// Dimensions returns the highest used row and column (1-based; 0 when empty).
// For loaded files without a trustworthy stored dimension the sheet is
// scanned once and the result cached; writes keep it current afterwards.
func (w *Workbook) Dimensions(sheet string) (maxRow, maxCol int, err error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return 0, 0, errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return 0, 0, err
	}
	if !st.dimsKnown {
		if err := w.scanDims(sheet, st); err != nil {
			return 0, 0, err
		}
	}
	return st.maxRow, st.maxCol, nil
}

func (w *Workbook) scanDims(sheet string, st *sheetState) error {
	rows, err := w.f.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()
	row := 0
	for rows.Next() {
		row++
		cols, err := rows.Columns(excelize.Options{RawCellValue: true})
		if err != nil {
			return err
		}
		if len(cols) > st.maxCol {
			st.maxCol = len(cols)
		}
		if len(cols) > 0 {
			st.maxRow = row
		}
	}
	st.dimsKnown = true
	return rows.Error()
}

// --- styling / structure ------------------------------------------------------

// SetNumberFormat applies a number-format code to a cell range through the
// style log, so it streams like any other style (structure.go).
func (w *Workbook) SetNumberFormat(sheet, ref, code string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	spec := compat.StyleSpec{"numberFormat": map[string]any{"formatCode": code}}
	return w.applySpecLocked(sheet, st, ref, spec)
}

// MergeCells merges a range like "A1:C3"; streaming sheets use the
// StreamWriter's native merge support instead of degrading.
func (w *Workbook) MergeCells(sheet, ref string) error {
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
	if st.sw != nil {
		return st.sw.MergeCell(tl, br)
	}
	if !st.random() {
		st.preMerges = append(st.preMerges, [2]string{tl, br})
		return nil
	}
	return w.f.MergeCell(sheet, tl, br)
}

func splitRange(ref string) (string, string, error) {
	parts := strings.Split(ref, ":")
	switch len(parts) {
	case 1:
		return parts[0], parts[0], nil
	case 2:
		return parts[0], parts[1], nil
	}
	return "", "", fmt.Errorf("easy-excel: invalid range %q", ref)
}

// --- save -------------------------------------------------------------------

// SaveXlsx flushes all stream writers and writes the workbook to path.
// The workbook stays usable afterwards (further writes use random mode).
func (w *Workbook) SaveXlsx(path string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	abs, err := w.policy.Resolve(path)
	if err != nil {
		return err
	}
	release, err := w.gate.AcquireHeavy()
	if err != nil {
		return err
	}
	defer release()
	if err := w.settleForSave(); err != nil {
		return err
	}
	return w.f.SaveAs(abs)
}

// WriteXlsxTo streams the workbook to an arbitrary writer (php:// targets).
func (w *Workbook) WriteXlsxTo(out io.Writer) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	release, err := w.gate.AcquireHeavy()
	if err != nil {
		return err
	}
	defer release()
	if err := w.settleForSave(); err != nil {
		return err
	}
	return w.f.Write(out)
}

// settleForSave brings the workbook into a saveable state: queued structure
// work that excelize cannot stream forces the documented one-time degrade,
// otherwise the stream writers are just flushed.
func (w *Workbook) settleForSave() error {
	if w.hasAnyPendingWork() {
		if err := w.ensureRandomAll(); err != nil {
			return err
		}
	}
	return w.flushStreams()
}

func (w *Workbook) flushStreams() error {
	for _, st := range w.sheets {
		if st.sw != nil {
			if err := st.sw.Flush(); err != nil {
				return err
			}
			st.sw = nil
		}
		// flushed sheets can no longer accept stream rows
		st.eligible = false
	}
	return nil
}

// CsvOptions mirrors PhpSpreadsheet's Csv writer settings (Phase-1 subset:
// enclosure is fixed to '"').
type CsvOptions struct {
	Delimiter    rune
	UseCRLF      bool
	UseBOM       bool
	GuardFormula bool // opt-in OWASP injection guard
}

// SaveCsv writes one sheet as CSV with formatted values.
func (w *Workbook) SaveCsv(path, sheet string, opts CsvOptions) error {
	w.mu.Lock()
	if w.closed {
		w.mu.Unlock()
		return errClosed
	}
	abs, err := w.policy.Resolve(path)
	if err != nil {
		w.mu.Unlock()
		return err
	}
	w.mu.Unlock()

	out, err := os.Create(abs)
	if err != nil {
		return err
	}
	defer out.Close()
	return w.writeCsv(out, sheet, opts)
}

// WriteCsvTo writes one sheet as CSV to an arbitrary writer.
func (w *Workbook) WriteCsvTo(out io.Writer, sheet string, opts CsvOptions) error {
	return w.writeCsv(out, sheet, opts)
}

func (w *Workbook) writeCsv(out io.Writer, sheet string, opts CsvOptions) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	release, err := w.gate.AcquireHeavy()
	if err != nil {
		return err
	}
	defer release()
	if err := w.ensureRandom(sheet, st); err != nil {
		return err
	}
	return writeCsvRows(w.f, out, sheet, opts)
}

// Close releases excelize resources and memory accounting; idempotent.
func (w *Workbook) Close() error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return nil
	}
	w.closed = true
	for _, st := range w.sheets {
		w.closeIter(st)
		st.sw = nil
	}
	w.gate.ReserveMemory(-w.estBytes)
	w.estBytes = 0
	return w.f.Close()
}
