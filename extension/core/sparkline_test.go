package core

import (
	"archive/zip"
	"io"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestAddSparklineGroupOnStreamedSheet(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "sparklines.xlsx")

	fillRows(t, w, "Worksheet", 1, 5)
	if err := w.AddSparklineGroup("Worksheet", `{"type":"column",
		"sparklines":[{"location":"G1","dataRange":"Worksheet!A1:B1"},{"location":"G2","dataRange":"Worksheet!A2:B2"}],
		"high":true,"colorSeries":"FF00B050"}`); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	assertSheetXMLContains(t, path, `type="column"`)
	assertSheetXMLContains(t, path, `<xm:f>Worksheet!A2:B2</xm:f><xm:sqref>G2</xm:sqref>`)
	assertSheetXMLContains(t, path, `high="true"`)
	assertSheetXMLContains(t, path, `rgb="FF00B050"`)

	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "B5"); v != "5" {
		t.Errorf("B5 = %q, want 5 — data must survive the sparkline flush", v)
	}
}

// Two groups on one sheet land in the same x14:sparklineGroups container
// rather than one replacing the other.
func TestAddSparklineGroupAppends(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "two-groups.xlsx")

	fillRows(t, w, "Worksheet", 1, 3)
	for _, spec := range []string{
		`{"type":"line","sparklines":[{"location":"G1","dataRange":"Worksheet!A1:B1"}]}`,
		`{"type":"stacked","sparklines":[{"location":"G2","dataRange":"Worksheet!A2:B2"}]}`,
	} {
		if err := w.AddSparklineGroup("Worksheet", spec); err != nil {
			t.Fatal(err)
		}
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	xml := worksheetXML(t, path, "xl/worksheets/sheet1.xml")
	if n := strings.Count(xml, "<x14:sparklineGroup "); n != 2 {
		t.Errorf("sparkline groups = %d, want 2", n)
	}
	if strings.Count(xml, "<x14:sparklineGroups") != 1 {
		t.Error("groups must share one x14:sparklineGroups container")
	}
	assertSheetXMLContains(t, path, `type="stacked"`)
}

func TestAddSparklineGroupRejectsBadSpec(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()

	if err := w.AddSparklineGroup("Worksheet", `{"type":"pie","sparklines":[{"location":"A1","dataRange":"B1:C1"}]}`); err == nil {
		t.Error("unknown type must fail at the call site")
	}
	if err := w.AddSparklineGroup("Nope", `{"sparklines":[{"location":"A1","dataRange":"B1:C1"}]}`); err == nil {
		t.Error("unknown sheet must fail")
	}
}

func worksheetXML(t *testing.T, path, part string) string {
	t.Helper()
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	for _, zf := range zr.File {
		if zf.Name != part {
			continue
		}
		r, err := zf.Open()
		if err != nil {
			t.Fatal(err)
		}
		defer r.Close()
		data, err := io.ReadAll(r)
		if err != nil {
			t.Fatal(err)
		}
		return string(data)
	}
	t.Fatalf("%s missing from %s", part, path)
	return ""
}

func TestLoadedSparklineXML(t *testing.T) {
	path := filepath.Join(t.TempDir(), "source.xlsx")
	src := excelize.NewFile()
	if _, err := src.NewSheet("Plain"); err != nil {
		t.Fatal(err)
	}
	if err := src.SetSheetRow("Sheet1", "A1", &[]any{1, 2, 3}); err != nil {
		t.Fatal(err)
	}
	if err := src.AddSparkline("Sheet1", &excelize.SparklineOptions{
		Location: []string{"D1"}, Range: []string{"Sheet1!A1:C1"}, Type: "column", High: true,
	}); err != nil {
		t.Fatal(err)
	}
	if err := src.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	src.Close()

	w, err := Open(path, "", testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()

	ext, err := w.LoadedSparklineXML("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{extURISparklineGroups, `<x14:sparklineGroup `, `type="column"`, `<xm:sqref>D1</xm:sqref>`} {
		if !strings.Contains(ext, want) {
			t.Errorf("loaded ext lacks %q:\n%s", want, ext)
		}
	}
	if !strings.HasPrefix(ext, "<ext ") || !strings.HasSuffix(ext, "</ext>") {
		t.Errorf("not a whole <ext> element: %.60s…%s", ext, ext[max(0, len(ext)-10):])
	}
	if ext, _ := w.LoadedSparklineXML("Plain"); ext != "" {
		t.Errorf("sheet without sparklines = %q", ext)
	}

	if err := w.RenameSheet("Sheet1", "Trend"); err != nil {
		t.Fatal(err)
	}
	if ext, _ := w.LoadedSparklineXML("Trend"); !strings.Contains(ext, "<x14:sparklineGroup ") {
		t.Error("loaded sparklines must follow a rename")
	}
	if _, err := w.LoadedSparklineXML("Sheet1"); err == nil {
		t.Error("old name must be gone after the rename")
	}
}

func TestLoadedSparklineXMLEmptyForNewWorkbook(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 2)
	if err := w.AddSparklineGroup("Worksheet", `{"sparklines":[{"location":"G1","dataRange":"Worksheet!A1:B1"}]}`); err != nil {
		t.Fatal(err)
	}
	if ext, err := w.LoadedSparklineXML("Worksheet"); err != nil || ext != "" {
		t.Errorf("new workbook = %q, %v; groups added in memory are not 'loaded'", ext, err)
	}
	if w.Degraded() {
		t.Error("asking for loaded sparklines must not end streaming")
	}
}

func TestExtElement(t *testing.T) {
	doc := []byte(`<worksheet><x:extLst><x:ext uri="{A}"><a/></x:ext><x:ext uri="` + extURISparklineGroups + `"><g/></x:ext></x:extLst></worksheet>`)
	if got := extElement(doc, extURISparklineGroups); got != `<x:ext uri="`+extURISparklineGroups+`"><g/></x:ext>` {
		t.Errorf("prefixed ext = %q", got)
	}
	if got := extElement([]byte(`<worksheet/>`), extURISparklineGroups); got != "" {
		t.Errorf("absent = %q", got)
	}
	if got := extElement([]byte(`<foo bar="`+extURISparklineGroups+`"/>`), extURISparklineGroups); got != "" {
		t.Errorf("uri outside an ext element = %q", got)
	}
}

func TestWithRootNamespaces(t *testing.T) {
	part := []byte(`<?xml version="1.0"?><worksheet xmlns="main" xmlns:x14="u14" xmlns:xr2="rev2&amp;x"><sheetData/></worksheet>`)
	got := withRootNamespaces(`<ext xmlns:x14="u14" uri="{U}"><x14:g xr2:uid="1"/></ext>`, part)
	want := `<ext xmlns:xr2="rev2&amp;x" xmlns:x14="u14" uri="{U}"><x14:g xr2:uid="1"/></ext>`
	if got != want {
		t.Errorf("got  %s\nwant %s", got, want)
	}
	if same := withRootNamespaces(`<ext uri="{U}"/>`, []byte("not xml <")); same != `<ext uri="{U}"/>` {
		t.Errorf("an unreadable part must leave the element alone, got %s", same)
	}
}

func TestCopySheetCarriesLoadedSparklines(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 3)
	if err := w.AddSparklineGroup("Worksheet", `{"type":"line","sparklines":[{"location":"E2","dataRange":"Worksheet!A2:C2"}]}`); err != nil {
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
	if _, err := r.CopySheetTo("Worksheet", "Copy"); err != nil {
		t.Fatal(err)
	}
	if ext, _ := r.LoadedSparklineXML("Copy"); !strings.Contains(ext, "sparklineGroup") {
		t.Error("the copy carries the source's sparklines in the file, so it must report them")
	}
}
