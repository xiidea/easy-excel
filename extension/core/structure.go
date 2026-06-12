package core

import (
	"fmt"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"

	"github.com/ronisaha/easy-excel/extension/compat"
)

// Phase-2 structure & styling (PLAN.md §5). Ops are recorded in per-sheet
// logs so the streaming write path survives the common report pattern
// (style the header, set widths, freeze panes, then bulk-write rows):
//
//   - cell styles and row heights queued before their rows are written are
//     inlined into the StreamWriter rows (zero degrade);
//   - column widths, freeze panes and merges queued before streaming starts
//     are applied through the StreamWriter's native support;
//   - everything excelize cannot stream (auto-filter, auto-size, hyperlinks,
//     comments, page setup, styles on already-streamed cells) stays queued
//     and is replayed after the one-time degrade — triggered by the first
//     read, or deferred all the way to save.
//
// Replay is idempotent: entries are marked done and the log is kept so later
// styles merge over earlier ones (PhpSpreadsheet's partial-style layering).

type styleEntry struct {
	r1, c1, r2, c2 int // rect; r2 == 0 means a full column (unbounded rows)
	spec           compat.StyleSpec
	inline         bool // rows were still unwritten when queued → streamable
	done           bool
}

type heightEntry struct {
	height float64
	done   bool
}

type opKind byte

const (
	opAutoFilter opKind = iota
	opAutoSize
	opHyperlink
	opComment
	opPageSetup
)

type pendingOp struct {
	kind        opKind
	ref, s1, s2 string
	a, b, c     int
}

type colWidthOp struct {
	c1, c2 int
	width  float64
}

// rectOf parses "A1", "A1:C3", "C" or "C:E" into a normalized rect.
// Full-column refs return r2 == 0.
func rectOf(ref string) (r1, c1, r2, c2 int, err error) {
	parts := strings.SplitN(ref, ":", 2)
	isCol := func(s string) bool {
		for _, ch := range s {
			if ch < 'A' || ch > 'Z' {
				return false
			}
		}
		return s != ""
	}
	if isCol(parts[0]) {
		c1, err = excelize.ColumnNameToNumber(parts[0])
		if err != nil {
			return
		}
		c2 = c1
		if len(parts) == 2 {
			if !isCol(parts[1]) {
				err = fmt.Errorf("easy-excel: invalid range %q", ref)
				return
			}
			if c2, err = excelize.ColumnNameToNumber(parts[1]); err != nil {
				return
			}
		}
		if c2 < c1 {
			c1, c2 = c2, c1
		}
		return 1, c1, 0, c2, nil
	}
	c1, r1, err = excelize.CellNameToCoordinates(parts[0])
	if err != nil {
		return
	}
	r2, c2 = r1, c1
	if len(parts) == 2 {
		if c2, r2, err = excelize.CellNameToCoordinates(parts[1]); err != nil {
			return
		}
	}
	if r2 < r1 {
		r1, r2 = r2, r1
	}
	if c2 < c1 {
		c1, c2 = c2, c1
	}
	return
}

func (e *styleEntry) containsRect(r1, c1, r2, c2 int) bool {
	if e.c1 > c1 || e.c2 < c2 {
		return false
	}
	if e.r2 == 0 {
		return e.r1 <= r1
	}
	return r2 != 0 && e.r1 <= r1 && e.r2 >= r2
}

func (e *styleEntry) containsCell(row, col int) bool {
	return e.c1 <= col && col <= e.c2 && e.r1 <= row && (e.r2 == 0 || row <= e.r2)
}

// random reports whether the sheet is past streaming (degraded or loaded).
func (st *sheetState) random() bool {
	return !st.eligible && st.sw == nil
}

func (st *sheetState) hasPendingWork() bool {
	if len(st.preWidths) > 0 || st.prePanes != nil || len(st.preMerges) > 0 || len(st.pending) > 0 {
		return true
	}
	for i := range st.styleLog {
		e := &st.styleLog[i]
		if e.done {
			continue
		}
		// full-column inline styles were applied to every streamed cell;
		// styling cells that were never written is skipped (COMPAT.md)
		if e.inline && e.r2 == 0 && st.lastRow > 0 {
			continue
		}
		return true
	}
	for _, h := range st.rowHeights {
		if !h.done {
			return true
		}
	}
	return false
}

// --- public ops ---------------------------------------------------------------

// ApplyStyle layers a PhpSpreadsheet style array (JSON) onto a range.
// Streaming sheets queue it (inlined if the rows are still unwritten);
// random-access sheets apply immediately.
func (w *Workbook) ApplyStyle(sheet, ref, jsonSpec string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	spec, err := compat.ParseStyleSpec(jsonSpec)
	if err != nil {
		return err
	}
	return w.applySpecLocked(sheet, st, ref, spec)
}

// applySpecLocked records (and, in random mode, applies) one style spec for a
// range; callers hold w.mu.
func (w *Workbook) applySpecLocked(sheet string, st *sheetState, ref string, spec compat.StyleSpec) error {
	r1, c1, r2, c2, err := rectOf(ref)
	if err != nil {
		return err
	}
	// validate eagerly so bad specs fail at the call site, not at save
	if _, err := compat.TranslateStyle(spec); err != nil {
		return err
	}
	e := styleEntry{r1: r1, c1: c1, r2: r2, c2: c2, spec: spec}
	if !st.random() {
		e.inline = r1 > st.lastRow
		st.styleLog = append(st.styleLog, e)
		return nil
	}
	st.styleLog = append(st.styleLog, e)
	return w.applyStyleEntry(sheet, st, len(st.styleLog)-1)
}

// SetColWidth sets the width of columns c1..c2 (1-based, inclusive).
func (w *Workbook) SetColWidth(sheet string, c1, c2 int, width float64) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if c2 < c1 {
		c1, c2 = c2, c1
	}
	if !st.random() {
		st.preWidths = append(st.preWidths, colWidthOp{c1: c1, c2: c2, width: width})
		return nil
	}
	return w.fSetColWidth(sheet, colWidthOp{c1: c1, c2: c2, width: width})
}

// SetColAutoSize marks columns for save-time width fitting (approximated by
// character count, see COMPAT.md).
func (w *Workbook) SetColAutoSize(sheet string, c1, c2 int) error {
	return w.queueOp(sheet, pendingOp{kind: opAutoSize, a: c1, b: c2})
}

// SetRowHeight sets one row's height; streamed-in-order rows carry it inline.
func (w *Workbook) SetRowHeight(sheet string, row int, height float64) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if st.random() {
		return w.f.SetRowHeight(sheet, row, height)
	}
	if st.rowHeights == nil {
		st.rowHeights = make(map[int]*heightEntry)
	}
	st.rowHeights[row] = &heightEntry{height: height}
	if row <= st.lastRow {
		// row already streamed; replayed after the save-time degrade
		return nil
	}
	return nil
}

// FreezePanes freezes rows above / columns left of topLeft ("" unfreezes),
// matching Worksheet::freezePane().
func (w *Workbook) FreezePanes(sheet, topLeft string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	panes, err := panesFor(topLeft)
	if err != nil {
		return err
	}
	if !st.random() {
		st.prePanes = panes
		return nil
	}
	return w.f.SetPanes(sheet, panes)
}

func panesFor(topLeft string) (*excelize.Panes, error) {
	if topLeft == "" || topLeft == "A1" {
		return &excelize.Panes{Freeze: false, Split: false}, nil
	}
	col, row, err := excelize.CellNameToCoordinates(topLeft)
	if err != nil {
		return nil, err
	}
	active := "bottomRight"
	switch {
	case col == 1:
		active = "bottomLeft"
	case row == 1:
		active = "topRight"
	}
	return &excelize.Panes{
		Freeze:      true,
		XSplit:      col - 1,
		YSplit:      row - 1,
		TopLeftCell: topLeft,
		ActivePane:  active,
		Selection:   []excelize.Selection{{SQRef: topLeft, ActiveCell: topLeft, Pane: active}},
	}, nil
}

// AutoFilter records a save-time auto-filter for the range.
func (w *Workbook) AutoFilter(sheet, ref string) error {
	return w.queueOp(sheet, pendingOp{kind: opAutoFilter, ref: ref})
}

// SetHyperlink attaches an external (or "sheet://" internal) link to a cell.
func (w *Workbook) SetHyperlink(sheet, cell, url, tooltip string) error {
	return w.queueOp(sheet, pendingOp{kind: opHyperlink, ref: cell, s1: url, s2: tooltip})
}

// SetComment sets (replacing any previous) a plain-text comment on a cell.
func (w *Workbook) SetComment(sheet, cell, author, text string) error {
	return w.queueOp(sheet, pendingOp{kind: opComment, ref: cell, s1: author, s2: text})
}

// SetPageSetup applies print layout; empty / -1 arguments are left unchanged.
func (w *Workbook) SetPageSetup(sheet, orientation string, paperSize, fitToWidth, fitToHeight int) error {
	return w.queueOp(sheet, pendingOp{
		kind: opPageSetup, s1: orientation, a: paperSize, b: fitToWidth, c: fitToHeight,
	})
}

// SetDefinedName registers a workbook- or sheet-scoped defined name; this is
// workbook-level XML, safe in any mode.
func (w *Workbook) SetDefinedName(name, refersTo, scopeSheet string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	dn := excelize.DefinedName{Name: name, RefersTo: refersTo, Scope: scopeSheet}
	if err := w.f.SetDefinedName(&dn); err != nil {
		// match PhpSpreadsheet: re-adding a name replaces it
		if delErr := w.f.DeleteDefinedName(&dn); delErr != nil {
			return err
		}
		return w.f.SetDefinedName(&dn)
	}
	return nil
}

func (w *Workbook) queueOp(sheet string, op pendingOp) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	if st.random() {
		return w.applyOp(sheet, op)
	}
	st.pending = append(st.pending, op)
	return nil
}

// --- stream-side application ---------------------------------------------------

// applyPreOps pushes queued widths, panes and merges through a fresh
// StreamWriter (all of which it only accepts before the first row).
func applyPreOps(sw *excelize.StreamWriter, st *sheetState) error {
	for _, cw := range st.preWidths {
		if err := sw.SetColWidth(cw.c1, cw.c2, cw.width); err != nil {
			return err
		}
	}
	st.preWidths = nil
	if st.prePanes != nil {
		if err := sw.SetPanes(st.prePanes); err != nil {
			return err
		}
		st.prePanes = nil
	}
	for _, m := range st.preMerges {
		if err := sw.MergeCell(m[0], m[1]); err != nil {
			return err
		}
	}
	st.preMerges = nil
	return nil
}

// inlineStyler resolves the interned style ID for cells of streamed rows from
// the queued inline entries, memoizing per distinct entry combination.
type inlineStyler struct {
	w       *Workbook
	st      *sheetState
	indices []int
	cache   map[string]int
}

func (w *Workbook) newInlineStyler(st *sheetState) *inlineStyler {
	var idx []int
	for i := range st.styleLog {
		if e := &st.styleLog[i]; e.inline && !e.done {
			idx = append(idx, i)
		}
	}
	if len(idx) == 0 {
		return nil
	}
	return &inlineStyler{w: w, st: st, indices: idx, cache: map[string]int{}}
}

func (s *inlineStyler) styleID(row, col int) (int, error) {
	var key strings.Builder
	var matched []int
	for _, i := range s.indices {
		if s.st.styleLog[i].containsCell(row, col) {
			matched = append(matched, i)
			fmt.Fprintf(&key, "%d,", i)
		}
	}
	if len(matched) == 0 {
		return 0, nil
	}
	if id, ok := s.cache[key.String()]; ok {
		return id, nil
	}
	merged := compat.StyleSpec{}
	for _, i := range matched {
		merged = compat.MergeSpec(merged, s.st.styleLog[i].spec)
	}
	id, err := s.w.styles.specID(s.w.f, merged)
	if err != nil {
		return 0, err
	}
	s.cache[key.String()] = id
	return id, nil
}

// markStreamed flags inline entries fully covered by streamed rows as done.
func (st *sheetState) markStreamed() {
	for i := range st.styleLog {
		e := &st.styleLog[i]
		if e.inline && !e.done && e.r2 != 0 && e.r2 <= st.lastRow {
			e.done = true
		}
	}
}

// --- replay (random-access application) -------------------------------------------

// replayAll applies every sheet's queued work; only valid once the whole
// workbook is in random-access mode (called from ensureRandom).
func (w *Workbook) replayAll() error {
	for name, st := range w.sheets {
		if err := w.replaySheet(name, st); err != nil {
			return err
		}
	}
	return nil
}

func (w *Workbook) replaySheet(sheet string, st *sheetState) error {
	for _, cw := range st.preWidths {
		if err := w.fSetColWidth(sheet, cw); err != nil {
			return err
		}
	}
	st.preWidths = nil
	if st.prePanes != nil {
		if err := w.f.SetPanes(sheet, st.prePanes); err != nil {
			return err
		}
		st.prePanes = nil
	}
	for _, m := range st.preMerges {
		if err := w.f.MergeCell(sheet, m[0], m[1]); err != nil {
			return err
		}
	}
	st.preMerges = nil
	for i := range st.styleLog {
		e := &st.styleLog[i]
		if e.done {
			continue
		}
		if e.inline && e.r2 == 0 && st.lastRow > 0 {
			e.done = true // streamed cells carry the style already
			continue
		}
		if err := w.applyStyleEntry(sheet, st, i); err != nil {
			return err
		}
	}
	for row, h := range st.rowHeights {
		if h.done {
			continue
		}
		if err := w.f.SetRowHeight(sheet, row, h.height); err != nil {
			return err
		}
		h.done = true
	}
	ops := st.pending
	st.pending = nil
	for _, op := range ops {
		if err := w.applyOp(sheet, op); err != nil {
			return err
		}
	}
	return nil
}

// applyStyleEntry interns the merge-folded spec for log[idx] and applies it.
// The fold layers every earlier entry whose rect contains this one, mirroring
// PhpSpreadsheet's partial-style application order.
func (w *Workbook) applyStyleEntry(sheet string, st *sheetState, idx int) error {
	e := &st.styleLog[idx]
	merged := compat.StyleSpec{}
	for j := 0; j <= idx; j++ {
		if p := &st.styleLog[j]; j == idx || p.containsRect(e.r1, e.c1, e.r2, e.c2) {
			merged = compat.MergeSpec(merged, p.spec)
		}
	}
	id, err := w.styles.specID(w.f, merged)
	if err != nil {
		return err
	}
	e.done = true
	if e.r2 == 0 {
		c1, _ := excelize.ColumnNumberToName(e.c1)
		c2, _ := excelize.ColumnNumberToName(e.c2)
		return w.f.SetColStyle(sheet, c1+":"+c2, id)
	}
	tl, err := excelize.CoordinatesToCellName(e.c1, e.r1)
	if err != nil {
		return err
	}
	br, err := excelize.CoordinatesToCellName(e.c2, e.r2)
	if err != nil {
		return err
	}
	return w.f.SetCellStyle(sheet, tl, br, id)
}

func (w *Workbook) fSetColWidth(sheet string, cw colWidthOp) error {
	c1, err := excelize.ColumnNumberToName(cw.c1)
	if err != nil {
		return err
	}
	c2, err := excelize.ColumnNumberToName(cw.c2)
	if err != nil {
		return err
	}
	return w.f.SetColWidth(sheet, c1, c2, cw.width)
}

func (w *Workbook) applyOp(sheet string, op pendingOp) error {
	switch op.kind {
	case opAutoFilter:
		tl, br, err := splitRange(op.ref)
		if err != nil {
			return err
		}
		return w.f.AutoFilter(sheet, tl+":"+br, nil)
	case opAutoSize:
		return w.autoSizeCols(sheet, op.a, op.b)
	case opHyperlink:
		link, linkType := op.s1, "External"
		if strings.HasPrefix(link, "sheet://") {
			link, linkType = strings.TrimPrefix(link, "sheet://"), "Location"
		}
		var opts []excelize.HyperlinkOpts
		if op.s2 != "" {
			tip := op.s2
			opts = append(opts, excelize.HyperlinkOpts{Tooltip: &tip})
		}
		return w.f.SetCellHyperLink(sheet, op.ref, link, linkType, opts...)
	case opComment:
		_ = w.f.DeleteComment(sheet, op.ref) // replace semantics
		return w.f.AddComment(sheet, excelize.Comment{
			Cell:      op.ref,
			Author:    op.s1,
			Paragraph: []excelize.RichTextRun{{Text: op.s2}},
		})
	case opPageSetup:
		layout := excelize.PageLayoutOptions{}
		if op.s1 != "" {
			o := op.s1
			layout.Orientation = &o
		}
		if op.a > 0 {
			ps := op.a
			layout.Size = &ps
		}
		if op.b >= 0 {
			fw := op.b
			layout.FitToWidth = &fw
		}
		if op.c >= 0 {
			fh := op.c
			layout.FitToHeight = &fh
		}
		return w.f.SetPageLayout(sheet, &layout)
	}
	return fmt.Errorf("easy-excel: unknown pending op %d", op.kind)
}

// autoSizeCols approximates PhpSpreadsheet's auto-size: widest formatted
// value per column (in characters) plus padding, capped at Excel's maximum.
func (w *Workbook) autoSizeCols(sheet string, c1, c2 int) error {
	rows, err := w.f.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()
	widest := make([]int, c2-c1+1)
	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return err
		}
		for c := c1; c <= c2 && c <= len(cols); c++ {
			if n := utf8.RuneCountInString(cols[c-1]); n > widest[c-c1] {
				widest[c-c1] = n
			}
		}
	}
	if err := rows.Error(); err != nil {
		return err
	}
	for i, n := range widest {
		if n == 0 {
			continue
		}
		width := float64(n) + 2
		if width > 254 {
			width = 254
		}
		if err := w.fSetColWidth(sheet, colWidthOp{c1: c1 + i, c2: c1 + i, width: width}); err != nil {
			return err
		}
	}
	return nil
}

// hasAnyPendingWork reports whether any sheet still has queued structure work
// (drives the save-time degrade decision).
func (w *Workbook) hasAnyPendingWork() bool {
	for _, st := range w.sheets {
		if st.hasPendingWork() {
			return true
		}
	}
	return false
}
