package core

import (
	"archive/zip"
	"io"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/ronisaha/easy-excel/extension/compat"
)

func fillRows(t *testing.T, w *Workbook, sheet string, from, to int) {
	t.Helper()
	rows := make([][]compat.Cell, 0, to-from+1)
	for r := from; r <= to; r++ {
		rows = append(rows, mustCells(t, "a", int64(r), 1.5))
	}
	if err := w.WriteRows(sheet, from, 1, rows); err != nil {
		t.Fatal(err)
	}
}

func reopen(t *testing.T, path string) *excelize.File {
	t.Helper()
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { f.Close() })
	return f
}

func cellStyle(t *testing.T, f *excelize.File, sheet, cell string) *excelize.Style {
	t.Helper()
	id, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		t.Fatal(err)
	}
	style, err := f.GetStyle(id)
	if err != nil {
		t.Fatal(err)
	}
	return style
}

func TestStyledHeaderStaysStreaming(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "styled.xlsx")

	// style + widths + panes + height queued before rows: the report pattern
	if err := w.ApplyStyle("Worksheet", "A1:C1",
		`{"font":{"bold":true},"fill":{"fillType":"solid","startColor":{"rgb":"FFFF00"}}}`); err != nil {
		t.Fatal(err)
	}
	if err := w.SetColWidth("Worksheet", 1, 2, 23.5); err != nil {
		t.Fatal(err)
	}
	if err := w.FreezePanes("Worksheet", "A2"); err != nil {
		t.Fatal(err)
	}
	if err := w.SetRowHeight("Worksheet", 1, 30); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 50)
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("styled header forced a degrade; should have streamed inline")
	}

	f := reopen(t, path)
	if s := cellStyle(t, f, "Worksheet", "B1"); s.Font == nil || !s.Font.Bold {
		t.Errorf("B1 should be bold: %+v", s.Font)
	} else if s.Fill.Pattern != 1 || len(s.Fill.Color) == 0 || s.Fill.Color[0] != "FFFF00" {
		t.Errorf("B1 fill wrong: %+v", s.Fill)
	}
	if s := cellStyle(t, f, "Worksheet", "B2"); s.Font != nil && s.Font.Bold {
		t.Error("B2 should not be bold")
	}
	if width, _ := f.GetColWidth("Worksheet", "B"); width != 23.5 {
		t.Errorf("col B width = %v, want 23.5", width)
	}
	if panes, _ := f.GetPanes("Worksheet"); !panes.Freeze || panes.YSplit != 1 {
		t.Errorf("panes wrong: %+v", panes)
	}
	if h, _ := f.GetRowHeight("Worksheet", 1); h != 30 {
		t.Errorf("row 1 height = %v, want 30", h)
	}
	// value integrity under styled streaming
	if v, _ := f.GetCellValue("Worksheet", "B50"); v != "50" {
		t.Errorf("B50 = %q, want 50", v)
	}
}

func TestNumberFormatStreamsInline(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "numfmt.xlsx")

	if err := w.SetNumberFormat("Worksheet", "C1:C10", "0.00%"); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("number format on future rows should not degrade")
	}
	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "C5"); v != "150.00%" {
		t.Errorf("C5 formatted = %q, want 150.00%%", v)
	}
}

func TestFullColumnFormatStreams(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "colfmt.xlsx")

	if err := w.SetNumberFormat("Worksheet", "B", "#,##0.00"); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 20)
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("full-column format should ride the stream")
	}
	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "B7"); v != "7.00" {
		t.Errorf("B7 formatted = %q, want 7.00", v)
	}
}

func TestStyleMergeLayersComponents(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "merge.xlsx")

	if err := w.ApplyStyle("Worksheet", "A1:C1", `{"font":{"bold":true}}`); err != nil {
		t.Fatal(err)
	}
	if err := w.ApplyStyle("Worksheet", "B1", `{"font":{"italic":true}}`); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 3)
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	if s := cellStyle(t, f, "Worksheet", "B1"); !s.Font.Bold || !s.Font.Italic {
		t.Errorf("B1 should layer bold+italic: %+v", s.Font)
	}
	if s := cellStyle(t, f, "Worksheet", "A1"); !s.Font.Bold || s.Font.Italic {
		t.Errorf("A1 should be bold only: %+v", s.Font)
	}
}

func TestStyleAfterWriteDegradesAtSaveOnly(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "late.xlsx")

	fillRows(t, w, "Worksheet", 1, 30)
	// rows already streamed: style queues, degrade deferred to save
	if err := w.ApplyStyle("Worksheet", "A1:A30", `{"font":{"bold":true}}`); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("styling already-written rows should defer the degrade to save")
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if !w.Degraded() {
		t.Fatal("save should have degraded to apply the late style")
	}
	f := reopen(t, path)
	if s := cellStyle(t, f, "Worksheet", "A15"); !s.Font.Bold {
		t.Error("late style lost")
	}
	if v, _ := f.GetCellValue("Worksheet", "B15"); v != "15" {
		t.Errorf("B15 = %q, want 15 (degrade must keep data)", v)
	}
}

func TestMergeCellsWhileStreaming(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "merged.xlsx")

	if err := w.MergeCells("Worksheet", "A1:C1"); err != nil { // before sw exists
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 5)
	if err := w.MergeCells("Worksheet", "A4:B5"); err != nil { // sw active
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("merges should not degrade a streaming sheet")
	}
	f := reopen(t, path)
	merged, err := f.GetMergeCells("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]bool{}
	for _, m := range merged {
		got[m.GetStartAxis()+":"+m.GetEndAxis()] = true
	}
	if !got["A1:C1"] || !got["A4:B5"] {
		t.Errorf("merges missing: %v", got)
	}
}

func TestDeferredOpsApplyAtSave(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "deferred.xlsx")

	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.AutoFilter("Worksheet", "A1:C10"); err != nil {
		t.Fatal(err)
	}
	if err := w.SetHyperlink("Worksheet", "A2", "https://example.com/x", "docs"); err != nil {
		t.Fatal(err)
	}
	if err := w.SetComment("Worksheet", "B3", "QA", "verify me"); err != nil {
		t.Fatal(err)
	}
	if err := w.SetComment("Worksheet", "B3", "QA", "verified"); err != nil { // replace
		t.Fatal(err)
	}
	if err := w.SetColAutoSize("Worksheet", 1, 1); err != nil {
		t.Fatal(err)
	}
	if err := w.SetPageSetup("Worksheet", "landscape", 9, 1, -1); err != nil {
		t.Fatal(err)
	}
	if err := w.SetDefinedName("data", "Worksheet!$A$1:$C$10", ""); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}

	f := reopen(t, path)
	if ok, link, _ := f.GetCellHyperLink("Worksheet", "A2"); !ok || link != "https://example.com/x" {
		t.Errorf("hyperlink: ok=%v link=%q", ok, link)
	}
	comments, err := f.GetComments("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	if len(comments) != 1 || comments[0].Author != "QA" || !strings.Contains(commentText(comments[0]), "verified") {
		t.Errorf("comment wrong: %+v", comments)
	}
	// col A's widest value is "a" (1 rune) → approximated width 1+2
	if width, _ := f.GetColWidth("Worksheet", "A"); width != 3 {
		t.Errorf("autosize width = %v, want 3", width)
	}
	layout, err := f.GetPageLayout("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	if layout.Orientation == nil || *layout.Orientation != "landscape" {
		t.Errorf("orientation: %+v", layout.Orientation)
	}
	names := f.GetDefinedName()
	found := false
	for _, n := range names {
		if n.Name == "data" && strings.Contains(n.RefersTo, "$A$1:$C$10") {
			found = true
		}
	}
	if !found {
		t.Errorf("defined name missing: %+v", names)
	}
	assertSheetXMLContains(t, path, `<autoFilter ref="$A$1:$C$10"`)
}

func commentText(c excelize.Comment) string {
	if c.Text != "" {
		return c.Text
	}
	var sb strings.Builder
	for _, p := range c.Paragraph {
		sb.WriteString(p.Text)
	}
	return sb.String()
}

// assertSheetXMLContains greps the first worksheet part of a saved file
// (excelize has no GetAutoFilter accessor).
func assertSheetXMLContains(t *testing.T, path, needle string) {
	t.Helper()
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	for _, zf := range zr.File {
		if !strings.HasPrefix(zf.Name, "xl/worksheets/") {
			continue
		}
		r, err := zf.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			t.Fatal(err)
		}
		if strings.Contains(string(data), needle) {
			return
		}
	}
	t.Errorf("no worksheet XML contains %q", needle)
}

func TestStructureOpsOnLoadedFile(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "src.xlsx")
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 5)
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	w.Close()

	// loaded files are random-access: ops apply immediately
	w2, err := Open(path, testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w2.Close()
	if err := w2.ApplyStyle("Worksheet", "A1", `{"font":{"bold":true}}`); err != nil {
		t.Fatal(err)
	}
	if err := w2.SetColWidth("Worksheet", 2, 2, 40); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(dir, "out.xlsx")
	if err := w2.SaveXlsx(out); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, out)
	if s := cellStyle(t, f, "Worksheet", "A1"); !s.Font.Bold {
		t.Error("style on loaded file lost")
	}
	if width, _ := f.GetColWidth("Worksheet", "B"); width != 40 {
		t.Errorf("width on loaded file = %v", width)
	}
}

func TestFormattedReadSeesQueuedStyles(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.SetNumberFormat("Worksheet", "C1:C5", "0.0"); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 5)
	// reading degrades and must replay the queued format first
	v, err := w.GetCell("Worksheet", "C2", GetFormatted)
	if err != nil {
		t.Fatal(err)
	}
	if v != "1.5" {
		t.Errorf("formatted C2 = %v, want 1.5", v)
	}
}
