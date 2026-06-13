package core

import (
	"path/filepath"
	"testing"
)

func TestInsertAndRemoveRows(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 5) // B column carries the row number

	if err := w.InsertRows("Worksheet", 2, 2); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Worksheet", "B4", GetFormatted); v != "2" {
		t.Errorf("after insert, B4 = %v, want 2 (shifted)", v)
	}
	maxRow, _, _ := w.Dimensions("Worksheet")
	if maxRow != 7 {
		t.Errorf("maxRow after insert = %d, want 7", maxRow)
	}

	if err := w.RemoveRows("Worksheet", 2, 2); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Worksheet", "B2", GetFormatted); v != "2" {
		t.Errorf("after remove, B2 = %v, want 2", v)
	}
	maxRow, _, _ = w.Dimensions("Worksheet")
	if maxRow != 5 {
		t.Errorf("maxRow after remove = %d, want 5", maxRow)
	}
}

func TestInsertAndRemoveCols(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)

	if err := w.InsertCols("Worksheet", 2, 1); err != nil { // before B
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Worksheet", "C2", GetFormatted); v != "2" {
		t.Errorf("after insert, C2 = %v, want 2 (shifted from B)", v)
	}
	if v, _ := w.GetCell("Worksheet", "B2", GetFormatted); v != "" {
		t.Errorf("inserted column should be empty, got %v", v)
	}
	if err := w.RemoveCols("Worksheet", 2, 1); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Worksheet", "B2", GetFormatted); v != "2" {
		t.Errorf("after remove, B2 = %v, want 2", v)
	}
	_, maxCol, _ := w.Dimensions("Worksheet")
	if maxCol != 3 {
		t.Errorf("maxCol = %d, want 3", maxCol)
	}
}

func TestInsertRowsMigratesPatchedFilter(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	dir := t.TempDir()
	first := filepath.Join(dir, "first.xlsx")
	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.AutoFilter("Worksheet", "A1:C10"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(first, ""); err != nil { // filter via container patch
		t.Fatal(err)
	}
	if err := w.InsertRows("Worksheet", 1, 1); err != nil { // must not strand the filter
		t.Fatal(err)
	}
	second := filepath.Join(dir, "second.xlsx")
	if err := w.SaveXlsx(second, ""); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, second, "autoFilter ref=")
}

func TestMoveAndCopySheet(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if _, err := w.AddSheet("Data"); err != nil {
		t.Fatal(err)
	}
	if _, err := w.AddSheet("Summary"); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Data", 1, 3)

	if err := w.MoveSheetTo("Summary", 0); err != nil {
		t.Fatal(err)
	}
	if got := w.Sheets(); got[0] != "Summary" {
		t.Errorf("order after move = %v", got)
	}
	if err := w.MoveSheetTo("Summary", 2); err != nil { // to the end
		t.Fatal(err)
	}
	if got := w.Sheets(); got[2] != "Summary" {
		t.Errorf("order after move-to-end = %v", got)
	}

	if _, err := w.CopySheetTo("Data", "Data Copy"); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Data Copy", "B2", GetFormatted); v != "2" {
		t.Errorf("copied sheet B2 = %v, want 2", v)
	}
	// the copy is independent
	if err := w.SetCell("Data Copy", "B2", mustCells(t, int64(99))[0]); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Data", "B2", GetFormatted); v != "2" {
		t.Errorf("source mutated by copy edit: %v", v)
	}
}

func TestWritesAfterSaveAreNotLost(t *testing.T) {
	// excelize silently discards model edits made after a StreamWriter
	// flush; the workbook must reopen before post-save mutations
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	dir := t.TempDir()
	fillRows(t, w, "Worksheet", 1, 5)
	if err := w.SaveXlsx(filepath.Join(dir, "first.xlsx"), ""); err != nil {
		t.Fatal(err)
	}

	if err := w.SetCell("Worksheet", "A1", mustCells(t, "EDITED")[0]); err != nil {
		t.Fatal(err)
	}
	if err := w.ApplyStyle("Worksheet", "A1", `{"font":{"bold":true}}`); err != nil {
		t.Fatal(err)
	}
	second := filepath.Join(dir, "second.xlsx")
	if err := w.SaveXlsx(second, ""); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, second)
	if v, _ := f.GetCellValue("Worksheet", "A1"); v != "EDITED" {
		t.Errorf("post-save edit lost: A1 = %q", v)
	}
	if s := cellStyle(t, f, "Worksheet", "A1"); s.Font == nil || !s.Font.Bold {
		t.Error("post-save style lost")
	}
	if v, _ := f.GetCellValue("Worksheet", "B5"); v != "5" {
		t.Errorf("streamed data lost across reopen: B5 = %q", v)
	}
}

func TestSheetViewHeaderFooterMargins(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)
	if err := w.SetSheetView("Worksheet",
		`{"showGridlines":false,"zoomScale":75,"rightToLeft":false,"tabColor":"FF0000"}`); err != nil {
		t.Fatal(err)
	}
	if err := w.SetHeaderFooter("Worksheet",
		`{"oddHeader":"&C&\"-,Bold\"Report","oddFooter":"&CPage &P of &N","differentFirst":false,"differentOddEven":false}`); err != nil {
		t.Fatal(err)
	}
	if err := w.SetPageMargins("Worksheet",
		`{"top":1.25,"bottom":1.25,"left":0.7,"right":0.7,"header":-1,"footer":-1}`); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "view.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	view, err := f.GetSheetView("Worksheet", 0)
	if err != nil {
		t.Fatal(err)
	}
	if view.ShowGridLines == nil || *view.ShowGridLines || view.ZoomScale == nil || *view.ZoomScale != 75 {
		t.Errorf("sheet view wrong: %+v", view)
	}
	props, err := f.GetSheetProps("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	if props.TabColorRGB == nil || *props.TabColorRGB != "FF0000" {
		t.Errorf("tab color wrong: %+v", props.TabColorRGB)
	}
	assertSheetXMLContains(t, path, "Page &amp;P of &amp;N")
	margins, err := f.GetPageMargins("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	if margins.Top == nil || *margins.Top != 1.25 || margins.Left == nil || *margins.Left != 0.7 {
		t.Errorf("margins wrong: top=%v left=%v", margins.Top, margins.Left)
	}
}
