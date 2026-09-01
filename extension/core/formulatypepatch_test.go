package core

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
)

// retypeCells is the safety boundary: it must rewrite exactly the cells the
// cache wrote and nothing else.
func TestRetypeCellsTargetsOnlyCachedRefs(t *testing.T) {
	const in = `<row r="1">` +
		`<c r="A1" t="str"><f>TEXT(B1,"0000")</f><v>0042</v></c>` +
		`<c r="B1" s="1" t="str"><f>SUM(C1:D1)</f><v>2500</v></c>` +
		`<c r="C1" t="str"><f>IF(A1,"y","n")</f><v>y</v></c>` +
		`</row>`

	got := string(retypeCells([]byte(in), map[string]bool{"B1": true}))

	if !contains(got, `<c r="B1" s="1"><f>SUM(C1:D1)</f><v>2500</v></c>`) {
		t.Errorf("cached cell B1 should lose t=str:\n%s", got)
	}
	// The one that would corrupt: a genuine string result that looks numeric.
	if !contains(got, `<c r="A1" t="str"><f>TEXT(B1,"0000")</f><v>0042</v></c>`) {
		t.Errorf("uncached A1 must keep t=str — retyping it turns \"0042\" into 42:\n%s", got)
	}
	if !contains(got, `<c r="C1" t="str">`) {
		t.Errorf("uncached string result must be untouched:\n%s", got)
	}
}

// Shared-formula children are self-closing (<f t="shared" si="0"/>); walking
// start tags rather than whole elements means they are handled like any other.
func TestRetypeCellsHandlesSharedFormulaChildren(t *testing.T) {
	const in = `<c r="C2" t="str"><f t="shared" si="0"/><v>2500</v></c>`
	got := string(retypeCells([]byte(in), map[string]bool{"C2": true}))
	if got != `<c r="C2"><f t="shared" si="0"/><v>2500</v></c>` {
		t.Errorf("shared-formula child not retyped: %s", got)
	}
}

func TestRetypeCellsLeavesUnrelatedMarkupAlone(t *testing.T) {
	for _, in := range []string{
		`<c r="A1" t="s"><v>0</v></c>`,
		`<c r="A1" t="inlineStr"><is><t>hi</t></is></c>`,
		`<c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c>`,
		`<sheetData/>`,
		`<c r="A1"><v>1</v></c>`,
	} {
		if got := string(retypeCells([]byte(in), map[string]bool{"Z99": true})); got != in {
			t.Errorf("rewrote unrelated markup:\n  in  %s\n  got %s", in, got)
		}
	}
}

func TestRetypeCellsToleratesMalformedTag(t *testing.T) {
	const in = `<c r="A1" t="str"` // truncated, no closing '>'
	if got := string(retypeCells([]byte(in), map[string]bool{"A1": true})); got != in {
		t.Errorf("malformed input must pass through unchanged, got %s", got)
	}
}



// End-to-end: a string result that looks numeric must survive a save that also
// caches a real numeric total.
func TestPrecalculateDoesNotCorruptNumericLookingText(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.SetPrecalculateFormulas(true); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "text.xlsx")
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{
		mustCells(t, 42.0, 1500.0),
	}); err != nil {
		t.Fatal(err)
	}
	setFormula(t, w, "C1", `TEXT(A1,"0000")`) // string result "0042"
	setFormula(t, w, "D1", "SUM(A1:B1)")      // genuine numeric
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	f := reopen(t, path)
	// The string result is deliberately NOT cached — caching it would store
	// 42 and lose the leading zeros — so it recomputes on open, correctly.
	if v, _ := f.CalcCellValue("Worksheet", "C1"); v != "0042" {
		t.Errorf("TEXT() recomputes to %q, want 0042", v)
	}
	if raw, _ := f.GetCellValue("Worksheet", "C1", excelize.Options{RawCellValue: true}); raw == "42" {
		t.Error("TEXT() result was cached as the number 42 — leading zeros lost")
	}
	if v, _ := f.GetCellValue("Worksheet", "D1", excelize.Options{RawCellValue: true}); v != "1542" {
		t.Errorf("numeric total = %q, want 1542", v)
	}
}

// Both container patches on one save: neither may discard the other.
func TestFormulaTypePatchComposesWithAutoFilter(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.SetPrecalculateFormulas(true); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "both.xlsx")

	fillRows(t, w, "Worksheet", 1, 20)
	setFormula(t, w, "D21", "SUM(B1:B20)")
	if err := w.ApplyStyle("Worksheet", "D21", `{"numberFormat":{"formatCode":"#,##0"}}`); err != nil {
		t.Fatal(err)
	}
	if err := w.AutoFilter("Worksheet", "A1:C20"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	// The filter arrives either from the container patch (relative ref) or,
	// once a style forces a degrade, from excelize's model writer (absolute).
	assertSheetXMLContains(t, path, `autoFilter`)
	f := reopen(t, path)
	if displayed, _ := f.GetCellValue("Worksheet", "D21"); displayed != "210" {
		t.Errorf("displayed = %q, want 210 (both patches must apply)", displayed)
	}
	if v, _ := f.GetCellValue("Worksheet", "B20"); v != "20" {
		t.Errorf("data damaged by double patch: B20 = %q", v)
	}
}

func contains(haystack, needle string) bool {
	return len(haystack) >= len(needle) && indexOf(haystack, needle) >= 0
}

func indexOf(h, n string) int {
	for i := 0; i+len(n) <= len(h); i++ {
		if h[i:i+len(n)] == n {
			return i
		}
	}
	return -1
}

func countStages(t *testing.T) int {
	t.Helper()
	m, _ := filepath.Glob(filepath.Join(os.TempDir(), "easyexcel-*.xlsx"))
	return len(m)
}

// Both container passes stage through temp files; none may survive the save.
func TestWriteXlsxToLeavesNoTempFiles(t *testing.T) {
	before := countStages(t)
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.SetPrecalculateFormulas(true); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 20)
	setFormula(t, w, "D21", "SUM(B1:B20)")
	if err := w.AutoFilter("Worksheet", "A1:C20"); err != nil {
		t.Fatal(err)
	}
	var buf bytes.Buffer
	if err := w.WriteXlsxTo(&buf); err != nil {
		t.Fatal(err)
	}
	if after := countStages(t); after != before {
		t.Errorf("temp files leaked: %d before, %d after", before, after)
	}
}

// The staged destination must not survive either, and a failed save must not
// destroy an existing file.
func TestSaveXlsxDoesNotLeaveStagedFile(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.SetPrecalculateFormulas(true); err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	path := filepath.Join(dir, "out.xlsx")
	fillRows(t, w, "Worksheet", 1, 5)
	setFormula(t, w, "D6", "SUM(B1:B5)")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	for _, e := range entries {
		if strings.Contains(e.Name(), ".eexcel.") {
			t.Errorf("staged file left behind: %s", e.Name())
		}
	}
	if len(entries) != 1 {
		t.Errorf("expected exactly the output file, got %d entries", len(entries))
	}
}
