package core

import (
	"archive/zip"
	"encoding/json"
	"fmt"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/xiidea/easy-excel/extension/compat"
)

// writeSales streams a header plus three data rows onto sheet.
func writeSales(t *testing.T, w *Workbook, sheet string) {
	t.Helper()
	rows := [][]any{
		{"Region", "Product", "Amount"},
		{"East", "Pens", int64(10)},
		{"West", "Pens", int64(20)},
		{"East", "Ink", int64(5)},
	}
	for i, r := range rows {
		if err := w.WriteRows(sheet, i+1, 1, [][]compat.Cell{mustCells(t, r...)}); err != nil {
			t.Fatal(err)
		}
	}
}

func pivotNames(t *testing.T, path, sheet string) []string {
	t.Helper()
	pts, err := reopen(t, path).GetPivotTables(sheet)
	if err != nil {
		t.Fatal(err)
	}
	names := make([]string, 0, len(pts))
	for _, p := range pts {
		names = append(names, p.Name)
	}
	sort.Strings(names)
	return names
}

const salesPivot = `{"name":"%s","sourceSheet":"Data","sourceRange":"A1:C4","location":"A3:A3",
	"rows":[{"field":"Region"}],"columns":[{"field":"Product"}],
	"data":[{"field":"Amount","subtotal":"sum","caption":"Sum of Amount"}]}`

func TestSetPivotTablesAcrossStreamedSheets(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	if _, err := w.AddSheet("Report"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	fillRows(t, w, "Report", 20, 22) // the target sheet streams too
	path := filepath.Join(t.TempDir(), "pivot.xlsx")

	if err := w.SetPivotTables("Report", "["+fmt.Sprintf(salesPivot, "Sales")+"]"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	pts, err := reopen(t, path).GetPivotTables("Report")
	if err != nil {
		t.Fatal(err)
	}
	if len(pts) != 1 {
		t.Fatalf("pivot tables on Report = %d, want 1", len(pts))
	}
	p := pts[0]
	if p.Name != "Sales" || p.DataRange != "Data!A1:C4" {
		t.Errorf("name/source = %q / %q", p.Name, p.DataRange)
	}
	if len(p.Rows) != 1 || p.Rows[0].Data != "Region" || len(p.Columns) != 1 || p.Columns[0].Data != "Product" {
		t.Errorf("axes = rows %+v cols %+v", p.Rows, p.Columns)
	}
	if len(p.Data) != 1 || p.Data[0].Data != "Amount" || p.Data[0].Name != "Sum of Amount" {
		t.Errorf("data = %+v", p.Data)
	}
	if v, _ := reopen(t, path).GetCellValue("Data", "C3"); v != "20" {
		t.Errorf("Data!C3 = %q — source data must survive", v)
	}
}

// PhpSpreadsheet regenerates pivots on every save: a second save replaces
// the set written by the first, and an empty list removes them.
func TestSetPivotTablesReplacesEarlierSave(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	dir := t.TempDir()

	save := func(name, specs string) []string {
		t.Helper()
		if err := w.SetPivotTables("Data", specs); err != nil {
			t.Fatal(err)
		}
		path := filepath.Join(dir, name)
		if err := w.SaveXlsx(path, ""); err != nil {
			t.Fatal(err)
		}
		return pivotNames(t, path, "Data")
	}

	two := "[" + fmt.Sprintf(salesPivot, "A") + "," + fmt.Sprintf(salesPivot, "B") + "]"
	if got := save("1.xlsx", two); len(got) != 2 || got[0] != "A" || got[1] != "B" {
		t.Fatalf("first save = %v", got)
	}
	if got := save("2.xlsx", "["+fmt.Sprintf(salesPivot, "C")+"]"); len(got) != 1 || got[0] != "C" {
		t.Fatalf("second save = %v, want only C", got)
	}
	if got := save("3.xlsx", "[]"); len(got) != 0 {
		t.Fatalf("empty list = %v, want none", got)
	}
}

// Saving the same set twice must not touch the package: excelize's delete
// leaves the old parts behind, so an unchanged resend is skipped.
func TestSetPivotTablesUnchangedResendIsNoop(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	spec := "[" + fmt.Sprintf(salesPivot, "A") + "]"
	path := filepath.Join(t.TempDir(), "twice.xlsx")
	for i := 0; i < 2; i++ {
		if err := w.SetPivotTables("Data", spec); err != nil {
			t.Fatal(err)
		}
		if err := w.SaveXlsx(path, ""); err != nil {
			t.Fatal(err)
		}
	}
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	parts := 0
	for _, zf := range zr.File {
		if strings.HasPrefix(zf.Name, "xl/pivotTables/pivotTable") {
			parts++
		}
	}
	if parts != 1 {
		t.Errorf("pivot table parts = %d, want 1 after an unchanged re-save", parts)
	}
}

func TestSetPivotTablesFollowsRename(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	dir := t.TempDir()
	if err := w.SetPivotTables("Data", "["+fmt.Sprintf(salesPivot, "A")+"]"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(filepath.Join(dir, "1.xlsx"), ""); err != nil {
		t.Fatal(err)
	}
	if err := w.RenameSheet("Data", "Facts"); err != nil {
		t.Fatal(err)
	}
	spec := `[{"name":"B","sourceSheet":"Facts","sourceRange":"A1:C4","location":"E3","data":[{"field":"Amount"}]}]`
	if err := w.SetPivotTables("Facts", spec); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "2.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	if got := pivotNames(t, path, "Facts"); len(got) != 1 || got[0] != "B" {
		t.Errorf("after rename = %v, want the old table replaced by B", got)
	}
}

func TestSetPivotTablesRejectsBadInput(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()

	for name, specs := range map[string]string{
		"not a list":   `{}`,
		"bad subtotal": `[{"name":"P","sourceSheet":"Worksheet","sourceRange":"A1:B5","location":"A3","data":[{"field":"x","subtotal":"median"}]}]`,
		"duplicate":    "[" + fmt.Sprintf(salesPivot, "P") + "," + fmt.Sprintf(salesPivot, "P") + "]",
	} {
		if err := w.SetPivotTables("Worksheet", specs); err == nil {
			t.Errorf("%s: expected an error", name)
		}
	}
	if err := w.SetPivotTables("Nope", "[]"); err == nil {
		t.Error("unknown sheet must fail")
	}
}

func TestLoadedPivotTablesJSON(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	if _, err := w.AddSheet("Report"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	if err := w.SetPivotTables("Report", "["+fmt.Sprintf(salesPivot, "Sales")+"]"); err != nil {
		t.Fatal(err)
	}
	if out, err := w.LoadedPivotTablesJSON("Report"); err != nil || out != "[]" {
		t.Errorf("in-memory workbook = %q, %v; generated tables are not 'loaded'", out, err)
	}
	path := filepath.Join(t.TempDir(), "source.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	w.Close()

	r, err := Open(path, "", testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()
	out, err := r.LoadedPivotTablesJSON("Report")
	if err != nil {
		t.Fatal(err)
	}
	var pivots []loadedPivot
	if err := json.Unmarshal([]byte(out), &pivots); err != nil {
		t.Fatal(err)
	}
	if len(pivots) != 1 {
		t.Fatalf("loaded pivots = %d, want 1", len(pivots))
	}
	if !strings.Contains(pivots[0].Table, `<pivotTableDefinition`) || !strings.Contains(pivots[0].Table, `name="Sales"`) {
		t.Errorf("table part = %.120s", pivots[0].Table)
	}
	if !strings.Contains(pivots[0].Cache, `<worksheetSource ref="A1:C4" sheet="Data"`) {
		t.Errorf("cache part = %.200s", pivots[0].Cache)
	}
	if out, _ := r.LoadedPivotTablesJSON("Data"); out != "[]" {
		t.Errorf("source sheet carries no pivots, got %q", out)
	}
	if err := r.RenameSheet("Report", "Summary"); err != nil {
		t.Fatal(err)
	}
	if out, _ := r.LoadedPivotTablesJSON("Summary"); !strings.Contains(out, "pivotTableDefinition") {
		t.Error("loaded pivots must follow a rename")
	}
}

func TestResolvePart(t *testing.T) {
	for _, c := range [][3]string{
		{"xl/worksheets/sheet2.xml", "../pivotTables/pivotTable1.xml", "xl/pivotTables/pivotTable1.xml"},
		{"xl/workbook.xml", "worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"},
		{"xl/workbook.xml", "/xl/worksheets/sheet1.xml", "xl/worksheets/sheet1.xml"},
	} {
		if got := resolvePart(c[0], c[1]); got != c[2] {
			t.Errorf("resolvePart(%q, %q) = %q, want %q", c[0], c[1], got, c[2])
		}
	}
}

// A set that fails part-way must stay tracked: the retry replaces the tables
// already added instead of duplicating them.
func TestSetPivotTablesPartialFailureIsRetried(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	if _, err := w.GetCell("Data", "A1", GetRaw); err != nil { // random mode: ops apply at once
		t.Fatal(err)
	}
	second := `{"name":"B","sourceSheet":"%s","sourceRange":"A1:C4","location":"H3","data":[{"field":"Amount"}]}`
	first := fmt.Sprintf(salesPivot, "A")
	if err := w.SetPivotTables("Data", "["+first+","+fmt.Sprintf(second, "Gone")+"]"); err == nil {
		t.Fatal("a missing source sheet must fail")
	}
	if err := w.SetPivotTables("Data", "["+first+","+fmt.Sprintf(second, "Data")+"]"); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "retry.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	if got := pivotNames(t, path, "Data"); len(got) != 2 || got[0] != "A" || got[1] != "B" {
		t.Errorf("after retry = %v, want [A B]", got)
	}
}

// Deletion is by name, so a generated table may not take the name of a table
// the loaded file carried: replacing it later would delete the file's own.
func TestSetPivotTablesRejectsNameOfLoadedTable(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	if err := w.RenameSheet("Worksheet", "Data"); err != nil {
		t.Fatal(err)
	}
	writeSales(t, w, "Data")
	spec := "[" + fmt.Sprintf(salesPivot, "PivotTable1") + "]"
	if err := w.SetPivotTables("Data", spec); err != nil {
		t.Fatal(err)
	}
	src := filepath.Join(t.TempDir(), "src.xlsx")
	if err := w.SaveXlsx(src, ""); err != nil {
		t.Fatal(err)
	}
	w.Close()

	r, err := Open(src, "", testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()
	if err := r.SetPivotTables("Data", spec); err == nil || !strings.Contains(err.Error(), "already exists") {
		t.Fatalf("err = %v, want a name clash", err)
	}
	out := filepath.Join(t.TempDir(), "out.xlsx")
	if err := r.SaveXlsx(out, ""); err != nil {
		t.Fatal(err)
	}
	if got := pivotNames(t, out, "Data"); len(got) != 1 {
		t.Errorf("tables = %v, want the file's own one only", got)
	}
}

func TestSetPivotTablesRejectsSingleCellSource(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	spec := `[{"name":"P","sourceSheet":"Worksheet","sourceRange":"$A$1","location":"H3","data":[{"field":"x"}]}]`
	if err := w.SetPivotTables("Worksheet", spec); err == nil {
		t.Error("a one-cell source range must fail at the call site")
	}
}
