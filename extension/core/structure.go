package core

import (
	"fmt"
	"sort"
	"strings"
	"unicode/utf8"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
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
	opValidation
	opConditional
	opImage
	opProtect
	opChart
	opUnmerge
	opSheetView
	opHeaderFooter
	opMargins
	opRichText
	opAutoFilterCols
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
	if len(st.preWidths) > 0 || st.prePanes != nil || len(st.preMerges) > 0 || len(st.pending) > 0 || st.preDefault {
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
	if err := w.mutable(); err != nil {
		return err
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
	if err := w.mutable(); err != nil {
		return err
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
		if err := w.mutable(); err != nil {
			return err
		}
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
	if err := w.mutable(); err != nil {
		return err
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
		if err := w.mutable(); err != nil {
			return err
		}
		return w.applyOp(sheet, op)
	}
	st.pending = append(st.pending, op)
	return nil
}

// --- stream-side application ---------------------------------------------------

// applyPreOps pushes queued widths, panes, merges and the workbook default
// style through a fresh StreamWriter (all only accepted before the first row).
func (w *Workbook) applyPreOps(sw *excelize.StreamWriter, st *sheetState) error {
	if st.preDefault && w.defaultSpec != nil {
		id, err := w.styles.specID(w.f, w.defaultSpec)
		if err != nil {
			return err
		}
		if err := sw.SetColStyle(1, 16384, id); err != nil {
			return err
		}
		st.preDefault = false
	}
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
// the queued inline entries. Matched entry sets are encoded as a bitmask and
// memoized, and entries whose row span covers the whole batch are additionally
// cached per column — so the hot loop for the common case (header style +
// full-column formats) is two map hits, no allocations.
type inlineStyler struct {
	w       *Workbook
	st      *sheetState
	indices []int
	cache   map[uint64]int
	colIDs  map[int]int // per-column result when all entries are row-uniform
	uniform bool
}

func (w *Workbook) newInlineStyler(st *sheetState, batchMinRow, batchMaxRow int) *inlineStyler {
	var idx []int
	for i := range st.styleLog {
		if e := &st.styleLog[i]; e.inline && !e.done {
			idx = append(idx, i)
		}
	}
	if len(idx) == 0 || len(idx) > 64 {
		if len(idx) > 64 {
			// fall back to marking them non-inline; replayed at save
			for _, i := range idx {
				st.styleLog[i].inline = false
			}
		}
		return nil
	}
	s := &inlineStyler{w: w, st: st, indices: idx, cache: map[uint64]int{}, uniform: true}
	for _, i := range idx {
		e := &st.styleLog[i]
		inside := e.r1 <= batchMinRow && (e.r2 == 0 || e.r2 >= batchMaxRow)
		outside := e.r1 > batchMaxRow || (e.r2 != 0 && e.r2 < batchMinRow)
		if !inside && !outside {
			s.uniform = false
			break
		}
	}
	if s.uniform {
		s.colIDs = map[int]int{}
	}
	return s
}

func (s *inlineStyler) styleID(row, col int) (int, error) {
	if s.uniform {
		if id, ok := s.colIDs[col]; ok {
			return id, nil
		}
	}
	var mask uint64
	for bit, i := range s.indices {
		if s.st.styleLog[i].containsCell(row, col) {
			mask |= 1 << uint(bit)
		}
	}
	id := 0
	if mask != 0 {
		if cached, ok := s.cache[mask]; ok {
			id = cached
		} else {
			merged := s.w.foldBase()
			for bit, i := range s.indices {
				if mask&(1<<uint(bit)) != 0 {
					merged = compat.MergeSpec(merged, s.st.styleLog[i].spec)
				}
			}
			interned, err := s.w.styles.specID(s.w.f, merged)
			if err != nil {
				return 0, err
			}
			s.cache[mask] = interned
			id = interned
		}
	}
	if s.uniform {
		s.colIDs[col] = id
	}
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
	if st.preDefault && w.defaultSpec != nil {
		if err := w.applyDefaultColStyle(sheet); err != nil {
			return err
		}
		st.preDefault = false
	}
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

// applyStyleEntry applies log[idx]: first its own rect with the fold of every
// earlier containing entry, then — because one SetCellStyle per rect would
// let a broad later style clobber narrower earlier ones (e.g. column number
// formats followed by a sheet-wide alignment) — the intersections with every
// earlier overlapping entry are re-applied with their own fold, restoring
// PhpSpreadsheet's per-cell layering for the overlap regions.
// exactStyleAreaLimit bounds the per-cell layering pass; bigger regions fall
// back to rect-based application (approximate for cross overlaps).
const exactStyleAreaLimit = 1 << 16

func (w *Workbook) applyStyleEntry(sheet string, st *sheetState, idx int) error {
	e := &st.styleLog[idx]
	e.done = true
	r1, c1, r2, c2 := e.r1, e.c1, e.r2, e.c2
	bound := r2
	if bound == 0 {
		// column default for rows written later; existing cells get the
		// exact pass below
		if err := w.applyFoldedRect(sheet, st, idx, r1, c1, 0, c2); err != nil {
			return err
		}
		bound = st.maxRow
		if bound < r1 {
			return nil
		}
	}
	if idx < 64 && (bound-r1+1)*(c2-c1+1) <= exactStyleAreaLimit {
		return w.applyStyleExact(sheet, st, idx, r1, c1, bound, c2)
	}
	return w.applyStyleRects(sheet, st, idx)
}

// applyStyleExact reproduces PhpSpreadsheet's per-cell layering for the rect:
// each cell's style is the fold of every entry up to idx containing it.
// Cells sharing an entry set are grouped into rectangular runs (consecutive
// identical rows × consecutive identical column masks), so the excelize call
// count stays proportional to the style structure, not the cell count.
func (w *Workbook) applyStyleExact(sheet string, st *sheetState, idx, r1, c1, r2, c2 int) error {
	width := c2 - c1 + 1
	rowMask := func(row int, out []uint64) {
		for i := range out {
			out[i] = 0
		}
		for j := 0; j <= idx; j++ {
			p := &st.styleLog[j]
			if p.r1 > row || (p.r2 != 0 && p.r2 < row) {
				continue
			}
			lo, hi := max(p.c1, c1), min(p.c2, c2)
			for c := lo; c <= hi; c++ {
				out[c-c1] |= 1 << uint(j)
			}
		}
	}
	cur := make([]uint64, width)
	prev := make([]uint64, width)
	groupStart := r1
	flush := func(rowEnd int, pattern []uint64) error {
		for s := 0; s < width; {
			m := pattern[s]
			t := s
			for t+1 < width && pattern[t+1] == m {
				t++
			}
			if m != 0 {
				if err := w.applyMaskRect(sheet, st, m, groupStart, c1+s, rowEnd, c1+t); err != nil {
					return err
				}
			}
			s = t + 1
		}
		return nil
	}
	for row := r1; row <= r2; row++ {
		rowMask(row, cur)
		if row > r1 && !slicesEqual(cur, prev) {
			if err := flush(row-1, prev); err != nil {
				return err
			}
			groupStart = row
		}
		copy(prev, cur)
	}
	return flush(r2, prev)
}

func slicesEqual(a, b []uint64) bool {
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func (w *Workbook) applyMaskRect(sheet string, st *sheetState, mask uint64, r1, c1, r2, c2 int) error {
	merged := w.foldBase()
	for j := 0; j < 64; j++ {
		if mask&(1<<uint(j)) != 0 {
			merged = compat.MergeSpec(merged, st.styleLog[j].spec)
		}
	}
	id, err := w.styles.specID(w.f, merged)
	if err != nil {
		return err
	}
	tl, err := excelize.CoordinatesToCellName(c1, r1)
	if err != nil {
		return err
	}
	br, err := excelize.CoordinatesToCellName(c2, r2)
	if err != nil {
		return err
	}
	return w.f.SetCellStyle(sheet, tl, br, id)
}

// applyStyleRects is the fallback for huge or >64-entry logs: the entry's
// rect plus its pairwise intersections with earlier overlapping entries,
// applied largest-first so nested intersections win. Cross overlaps between
// intersections remain approximate (COMPAT.md §12).
func (w *Workbook) applyStyleRects(sheet string, st *sheetState, idx int) error {
	e := &st.styleLog[idx]
	type rect struct{ r1, c1, r2, c2 int }
	seen := map[rect]bool{}
	rects := []rect{{e.r1, e.c1, e.r2, e.c2}}
	seen[rects[0]] = true
	for j := 0; j < idx; j++ {
		p := &st.styleLog[j]
		if p.containsRect(e.r1, e.c1, e.r2, e.c2) {
			continue // already part of e's fold
		}
		r1, c1, r2, c2, ok := intersectRects(p, e)
		if !ok {
			continue
		}
		if r := (rect{r1, c1, r2, c2}); !seen[r] {
			seen[r] = true
			rects = append(rects, r)
		}
	}
	area := func(r rect) int {
		if r.r2 == 0 {
			return int(^uint(0) >> 1)
		}
		return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1)
	}
	sort.SliceStable(rects, func(a, b int) bool { return area(rects[a]) > area(rects[b]) })
	for _, r := range rects {
		if err := w.applyFoldedRect(sheet, st, idx, r.r1, r.c1, r.r2, r.c2); err != nil {
			return err
		}
	}
	return nil
}

// intersectRects returns the overlap of two entries (r2 == 0 = unbounded rows).
func intersectRects(a, b *styleEntry) (r1, c1, r2, c2 int, ok bool) {
	c1 = max(a.c1, b.c1)
	c2 = min(a.c2, b.c2)
	if c1 > c2 {
		return 0, 0, 0, 0, false
	}
	r1 = max(a.r1, b.r1)
	switch {
	case a.r2 == 0 && b.r2 == 0:
		r2 = 0
	case a.r2 == 0:
		r2 = b.r2
	case b.r2 == 0:
		r2 = a.r2
	default:
		r2 = min(a.r2, b.r2)
	}
	if r2 != 0 && r1 > r2 {
		return 0, 0, 0, 0, false
	}
	return r1, c1, r2, c2, true
}

// applyFoldedRect sets the rect's style to the merge of every entry up to
// upTo (inclusive) that contains the rect.
func (w *Workbook) applyFoldedRect(sheet string, st *sheetState, upTo, r1, c1, r2, c2 int) error {
	merged := w.foldBase()
	for j := 0; j <= upTo; j++ {
		if p := &st.styleLog[j]; p.containsRect(r1, c1, r2, c2) {
			merged = compat.MergeSpec(merged, p.spec)
		}
	}
	id, err := w.styles.specID(w.f, merged)
	if err != nil {
		return err
	}
	if r2 == 0 {
		cn1, _ := excelize.ColumnNumberToName(c1)
		cn2, _ := excelize.ColumnNumberToName(c2)
		return w.f.SetColStyle(sheet, cn1+":"+cn2, id)
	}
	tl, err := excelize.CoordinatesToCellName(c1, r1)
	if err != nil {
		return err
	}
	br, err := excelize.CoordinatesToCellName(c2, r2)
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
	return w.applyOpPhase3(sheet, op)
}

// autoSizeCols approximates PhpSpreadsheet's auto-size: widest formatted
// value per column (in characters) plus padding, capped at Excel's maximum.
// Cells inside merged ranges are skipped, like PhpSpreadsheet — a merged
// title spanning the column must not stretch it.
func (w *Workbook) autoSizeCols(sheet string, c1, c2 int) error {
	merged, err := w.f.GetMergeCells(sheet)
	if err != nil {
		return err
	}
	type rect struct{ r1, c1, r2, c2 int }
	var merges []rect
	for _, m := range merged {
		mc1, mr1, err1 := excelize.CellNameToCoordinates(m.GetStartAxis())
		mc2, mr2, err2 := excelize.CellNameToCoordinates(m.GetEndAxis())
		if err1 == nil && err2 == nil {
			merges = append(merges, rect{mr1, mc1, mr2, mc2})
		}
	}
	inMerge := func(row, col int) bool {
		for _, m := range merges {
			if m.r1 <= row && row <= m.r2 && m.c1 <= col && col <= m.c2 {
				return true
			}
		}
		return false
	}
	rows, err := w.f.Rows(sheet)
	if err != nil {
		return err
	}
	defer rows.Close()
	widest := make([]int, c2-c1+1)
	rowNum := 0
	for rows.Next() {
		rowNum++
		cols, err := rows.Columns()
		if err != nil {
			return err
		}
		for c := c1; c <= c2 && c <= len(cols); c++ {
			if inMerge(rowNum, c) {
				continue
			}
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
