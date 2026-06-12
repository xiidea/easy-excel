package core

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/ronisaha/easy-excel/extension/compat"
	"github.com/ronisaha/easy-excel/extension/limits"
)

func testEnv() *Env {
	return &Env{Gate: limits.NewGate(limits.Config{})}
}

func mustCells(t *testing.T, vals ...any) []compat.Cell {
	t.Helper()
	out := make([]compat.Cell, len(vals))
	for i, v := range vals {
		c, err := compat.Decode(v)
		if err != nil {
			t.Fatal(err)
		}
		out[i] = c
	}
	return out
}

func TestNewWorkbookDefaultSheetName(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	sheets := w.Sheets()
	if len(sheets) != 1 || sheets[0] != "Worksheet" {
		t.Fatalf("PhpSpreadsheet default sheet must be \"Worksheet\", got %v", sheets)
	}
}

func TestStreamWriteSaveReload(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "out.xlsx")

	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	rows := [][]compat.Cell{
		mustCells(t, "name", "qty", "price"),
		mustCells(t, "widget", int64(3), 9.99),
		mustCells(t, "=B2*C2", true, "0042"),
	}
	if err := w.WriteRows("Worksheet", 1, 1, rows); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("sequential writes must stay in streaming mode")
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}

	r, err := Open(path, testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()
	if v, _ := r.GetCell("Worksheet", "A1", GetRaw); v != "name" {
		t.Fatalf("A1 = %#v", v)
	}
	if v, _ := r.GetCell("Worksheet", "B2", GetRaw); v != float64(3) {
		t.Fatalf("B2 = %#v, want float64(3)", v)
	}
	if v, _ := r.GetCell("Worksheet", "B3", GetRaw); v != true {
		t.Fatalf("B3 = %#v, want true", v)
	}
	// "0042" must survive as a string (leading-zero rule)
	if v, _ := r.GetCell("Worksheet", "C3", GetRaw); v != "0042" {
		t.Fatalf("C3 = %#v, want \"0042\"", v)
	}
	// formula cells round-trip as "=..." like Cell::getValue()
	if v, _ := r.GetCell("Worksheet", "A3", GetRaw); v != "=B2*C2" {
		t.Fatalf("A3 = %#v, want \"=B2*C2\"", v)
	}
	if v, err := r.GetCell("Worksheet", "A3", GetCalculated); err != nil || v != "29.97" {
		t.Fatalf("calc A3 = %#v, %v", v, err)
	}
	// regression: excelize writes a degenerate <dimension ref="A1"/>;
	// dimensions of a reloaded file must come from a lazy scan
	if mr, mc, err := r.Dimensions("Worksheet"); err != nil || mr != 3 || mc != 3 {
		t.Fatalf("reloaded dims = %d,%d, %v; want 3,3", mr, mc, err)
	}
}

func TestOutOfOrderWriteDegradesOnce(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{
		mustCells(t, "r1"), mustCells(t, "r2"), mustCells(t, "r3"),
	}); err != nil {
		t.Fatal(err)
	}
	// going back to row 2 forces the documented fallback
	if err := w.WriteRows("Worksheet", 2, 2, [][]compat.Cell{mustCells(t, "patched")}); err != nil {
		t.Fatal(err)
	}
	if !w.Degraded() {
		t.Fatal("out-of-order write must degrade to random mode")
	}
	// streamed data and the patch must both be present
	if v, _ := w.GetCell("Worksheet", "A1", GetRaw); v != "r1" {
		t.Fatalf("A1 = %#v", v)
	}
	if v, _ := w.GetCell("Worksheet", "A3", GetRaw); v != "r3" {
		t.Fatalf("A3 = %#v", v)
	}
	if v, _ := w.GetCell("Worksheet", "B2", GetRaw); v != "patched" {
		t.Fatalf("B2 = %#v", v)
	}
	// further writes keep working in random mode
	if err := w.WriteRows("Worksheet", 10, 1, [][]compat.Cell{mustCells(t, int64(7))}); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Worksheet", "A10", GetRaw); v != float64(7) {
		t.Fatalf("A10 = %#v", v)
	}
}

func TestReadOnStreamingSheetDegrades(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	_ = w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{mustCells(t, "x")})
	if v, err := w.GetCell("Worksheet", "A1", GetRaw); err != nil || v != "x" {
		t.Fatalf("read-back: %#v, %v", v, err)
	}
	if !w.Degraded() {
		t.Fatal("reading a streamed sheet must degrade")
	}
}

func TestExplicitMarkersThroughWriteRows(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	row := []compat.Cell{}
	for _, v := range []any{
		[]any{compat.MarkString, "=NOT_A_FORMULA()"},
		[]any{compat.MarkNumeric, "19.5"},
	} {
		c, err := compat.Decode(v)
		if err != nil {
			t.Fatal(err)
		}
		row = append(row, c)
	}
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{row}); err != nil {
		t.Fatal(err)
	}
	if v, _ := w.GetCell("Worksheet", "A1", GetRaw); v != "=NOT_A_FORMULA()" {
		// explicit strings are stored as text; getValue still shows the text
		t.Fatalf("A1 = %#v", v)
	}
	if f, _ := w.f.GetCellFormula("Worksheet", "A1"); f != "" {
		t.Fatalf("A1 must not be a formula cell, got %q", f)
	}
	if v, _ := w.GetCell("Worksheet", "B1", GetRaw); v != 19.5 {
		t.Fatalf("B1 = %#v", v)
	}
}

func TestChunkedSequentialRead(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	var rows [][]compat.Cell
	for i := 1; i <= 25; i++ {
		rows = append(rows, mustCells(t, int64(i), "row"))
	}
	_ = w.WriteRows("Worksheet", 1, 1, rows)

	got := 0
	for start := 1; ; {
		chunk, more, err := w.ReadRows("Worksheet", start, 10, true, false)
		if err != nil {
			t.Fatal(err)
		}
		got += len(chunk)
		start += len(chunk)
		if !more || len(chunk) == 0 {
			break
		}
	}
	if got != 25 {
		t.Fatalf("read %d rows, want 25", got)
	}
	// rewind: a lower start must transparently restart the iterator
	chunk, _, err := w.ReadRows("Worksheet", 1, 1, true, false)
	if err != nil || len(chunk) != 1 || chunk[0][0] != "1" {
		t.Fatalf("rewind read: %v, %v", chunk, err)
	}
}

// Documents how sparse sheets behave through the iterator: excelize must pad
// skipped rows, otherwise toArray() row alignment breaks (COMPAT.md).
func TestSparseRowAlignment(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	_ = w.SetCell("Worksheet", "A1", mustCells(t, "first")[0])
	_ = w.SetCell("Worksheet", "A5", mustCells(t, "fifth")[0])
	rows, _, err := w.ReadRows("Worksheet", 1, 10, true, false)
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 5 {
		t.Fatalf("expected 5 rows (gaps padded), got %d: %v", len(rows), rows)
	}
	if rows[0][0] != "first" || rows[4][0] != "fifth" {
		t.Fatalf("row alignment broken: %v", rows)
	}
	for i := 1; i <= 3; i++ {
		if len(rows[i]) != 0 {
			t.Fatalf("row %d should be empty, got %v", i+1, rows[i])
		}
	}
}

func TestDimensions(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	if r, c, _ := w.Dimensions("Worksheet"); r != 0 || c != 0 {
		t.Fatalf("empty sheet dims = %d,%d", r, c)
	}
	_ = w.WriteRows("Worksheet", 2, 3, [][]compat.Cell{mustCells(t, "a", "b"), mustCells(t, "c")})
	r, c, err := w.Dimensions("Worksheet")
	if err != nil || r != 3 || c != 4 {
		t.Fatalf("dims = %d,%d, %v; want 3,4", r, c, err)
	}
}

func TestSheetManagement(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	idx, err := w.AddSheet("Data")
	if err != nil || idx < 0 {
		t.Fatalf("AddSheet: %d, %v", idx, err)
	}
	if _, err := w.AddSheet("Data"); err == nil {
		t.Fatal("duplicate sheet must error")
	}
	if err := w.RenameSheet("Data", "Numbers"); err != nil {
		t.Fatal(err)
	}
	if got := w.Sheets(); len(got) != 2 || got[1] != "Numbers" {
		t.Fatalf("sheets = %v", got)
	}
	if err := w.SetActiveSheet(1); err != nil {
		t.Fatal(err)
	}
	if pos, name := w.ActiveSheet(); pos != 1 || name != "Numbers" {
		t.Fatalf("active = %d %q", pos, name)
	}
	// writes to the renamed sheet keep streaming state
	if err := w.WriteRows("Numbers", 1, 1, [][]compat.Cell{mustCells(t, int64(1))}); err != nil {
		t.Fatal(err)
	}
	if err := w.DeleteSheet("Numbers"); err != nil {
		t.Fatal(err)
	}
	if err := w.DeleteSheet("Worksheet"); err == nil {
		t.Fatal("removing the last sheet must error")
	}
}

func TestTwoSheetsStreamIndependently(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	if _, err := w.AddSheet("Second"); err != nil {
		t.Fatal(err)
	}
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{mustCells(t, "a")}); err != nil {
		t.Fatal(err)
	}
	if err := w.WriteRows("Second", 1, 1, [][]compat.Cell{mustCells(t, "b")}); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("independent sheet streams must not degrade")
	}
	var buf bytes.Buffer
	if err := w.WriteXlsxTo(&buf); err != nil {
		t.Fatal(err)
	}
	f, err := excelize.OpenReader(bytes.NewReader(buf.Bytes()))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	if v, _ := f.GetCellValue("Worksheet", "A1"); v != "a" {
		t.Fatalf("Worksheet!A1 = %q", v)
	}
	if v, _ := f.GetCellValue("Second", "A1"); v != "b" {
		t.Fatalf("Second!A1 = %q", v)
	}
}

func TestNumberFormatAndFormattedRead(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	_ = w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{mustCells(t, 0.125)})
	if err := w.SetNumberFormat("Worksheet", "A1", "0.00%"); err != nil {
		t.Fatal(err)
	}
	v, err := w.GetCell("Worksheet", "A1", GetFormatted)
	if err != nil || v != "12.50%" {
		t.Fatalf("formatted = %#v, %v", v, err)
	}
	// raw value untouched
	if v, _ := w.GetCell("Worksheet", "A1", GetRaw); v != 0.125 {
		t.Fatalf("raw = %#v", v)
	}
}

func TestMergeCells(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	_ = w.SetCell("Worksheet", "A1", mustCells(t, "title")[0])
	if err := w.MergeCells("Worksheet", "A1:C1"); err != nil {
		t.Fatal(err)
	}
}

func TestSaveCsv(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "out.csv")
	w, _ := New(testEnv())
	defer w.Close()
	_ = w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{
		mustCells(t, "a", "b,c", `say "hi"`),
		mustCells(t, "=1+1", int64(2), "-x"),
	})
	if err := w.SaveCsv(path, "Worksheet", CsvOptions{GuardFormula: true}); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	got := string(data)
	if !strings.Contains(got, `"b,c"`) || !strings.Contains(got, `"say ""hi"""`) {
		t.Fatalf("csv quoting wrong:\n%s", got)
	}
	// formula cells export their calculated value; the guard prefixes
	// dangerous leading characters
	if !strings.Contains(got, "'-x") {
		t.Fatalf("injection guard missing:\n%s", got)
	}
}

func TestPathPolicyEnforcedOnSave(t *testing.T) {
	w, _ := New(testEnv())
	defer w.Close()
	if err := w.SaveXlsx("https://evil/out.xlsx"); err == nil {
		t.Fatal("URL scheme must be rejected")
	}
}

func TestClosedWorkbookErrors(t *testing.T) {
	w, _ := New(testEnv())
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	if err := w.Close(); err != nil {
		t.Fatal("Close must be idempotent")
	}
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{mustCells(t, "x")}); err == nil {
		t.Fatal("write after close must error")
	}
}

func TestMemoryAccountingReleasedOnClose(t *testing.T) {
	gate := limits.NewGate(limits.Config{})
	env := &Env{Gate: gate}
	w, err := New(env)
	if err != nil {
		t.Fatal(err)
	}
	if gate.MemoryUsed() == 0 {
		t.Fatal("live workbook must be accounted")
	}
	_ = w.Close()
	if got := gate.MemoryUsed(); got != 0 {
		t.Fatalf("accounting leak after close: %d", got)
	}
}
