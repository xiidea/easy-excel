package core

import (
	"bytes"
	"image"
	"image/png"
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/ronisaha/easy-excel/extension/compat"
)

func TestAutoFilterPatchKeepsStreaming(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "filtered.xlsx")

	fillRows(t, w, "Worksheet", 1, 200)
	if err := w.AutoFilter("Worksheet", "A1:C200"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("auto-filter alone must not degrade: it is patched into the container")
	}
	assertSheetXMLContains(t, path, `<autoFilter ref="A1:C200"/>`)

	// the patched container must stay a valid workbook with intact data
	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "B200"); v != "200" {
		t.Errorf("B200 = %q after patch, want 200", v)
	}
}

func TestAutoFilterWithStyledStream(t *testing.T) {
	// header style (inlined) + auto-filter (patched): still no degrade
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "styled-filtered.xlsx")

	if err := w.ApplyStyle("Worksheet", "A1:C1", `{"font":{"bold":true}}`); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 50)
	if err := w.AutoFilter("Worksheet", "A1:C50"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if w.Degraded() {
		t.Fatal("styled header + auto-filter should stream end to end")
	}
	f := reopen(t, path)
	if s := cellStyle(t, f, "Worksheet", "A1"); !s.Font.Bold {
		t.Error("header style lost in patched file")
	}
	assertSheetXMLContains(t, path, `<autoFilter ref="A1:C50"/>`)
}

func TestAutoFilterWithOtherPendingUsesDegrade(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "mixed.xlsx")

	fillRows(t, w, "Worksheet", 1, 20)
	if err := w.AutoFilter("Worksheet", "A1:C20"); err != nil {
		t.Fatal(err)
	}
	if err := w.SetHyperlink("Worksheet", "A1", "https://example.com", ""); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	if !w.Degraded() {
		t.Fatal("hyperlink forces the degrade; the filter should ride it")
	}
	f := reopen(t, path)
	if ok, link, _ := f.GetCellHyperLink("Worksheet", "A1"); !ok || link != "https://example.com" {
		t.Errorf("hyperlink lost: %v %q", ok, link)
	}
	assertSheetXMLContains(t, path, `autoFilter ref=`)
}

func TestDataValidationRoundTrip(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "validated.xlsx")

	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.SetDataValidation("Worksheet", "C2:C10",
		`{"type":"list","formula1":"open,paid,void","allowBlank":true,"showErrorMessage":true,"errorTitle":"Bad status","error":"pick one"}`); err != nil {
		t.Fatal(err)
	}
	if err := w.SetDataValidation("Worksheet", "B2:B10",
		`{"type":"whole","operator":"between","formula1":"1","formula2":"100","showErrorMessage":true}`); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	dvs, err := f.GetDataValidations("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	if len(dvs) != 2 {
		t.Fatalf("want 2 validations, got %d", len(dvs))
	}
	byRef := map[string]*excelize.DataValidation{}
	for _, dv := range dvs {
		byRef[dv.Sqref] = dv
	}
	list := byRef["C2:C10"]
	if list == nil || list.Type != "list" || list.Formula1 != `"open,paid,void"` {
		t.Errorf("list validation wrong: %+v", list)
	}
	whole := byRef["B2:B10"]
	if whole == nil || whole.Type != "whole" || whole.Operator != "between" || whole.Formula1 != "1" || whole.Formula2 != "100" {
		t.Errorf("whole validation wrong: %+v", whole)
	}
}

func TestConditionalFormatRoundTrip(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "conditional.xlsx")

	fillRows(t, w, "Worksheet", 1, 10)
	rules := `[
		{"type":"cellIs","operator":"greaterThan","conditions":["5"],
		 "style":{"font":{"bold":true},"fill":{"fillType":"solid","startColor":{"rgb":"FFC7CE"}}}},
		{"type":"colorScale","colorScale":{"minColor":"FF0000","maxColor":"00FF00"}}
	]`
	if err := w.SetConditionalFormat("Worksheet", "B2:B10", rules); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	formats, err := f.GetConditionalFormats("Worksheet")
	if err != nil {
		t.Fatal(err)
	}
	opts, ok := formats["B2:B10"]
	if !ok || len(opts) != 2 {
		t.Fatalf("conditional formats missing: %+v", formats)
	}
	if opts[0].Type != "cell" || opts[0].Criteria == "" || opts[0].Format == nil {
		t.Errorf("cellIs rule wrong: %+v", opts[0])
	}
	if opts[1].Type != "2_color_scale" {
		t.Errorf("color scale rule wrong: %+v", opts[1])
	}
}

func TestAddImageScalesToRequestedSize(t *testing.T) {
	dir := t.TempDir()
	imgPath := filepath.Join(dir, "logo.png")
	img := image.NewRGBA(image.Rect(0, 0, 40, 20))
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(imgPath, buf.Bytes(), 0o644); err != nil {
		t.Fatal(err)
	}

	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)
	if err := w.AddImage("Worksheet", "E2", `{"path":"`+imgPath+`","width":80}`); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "with-image.xlsx")
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	pics, err := f.GetPictures("Worksheet", "E2")
	if err != nil {
		t.Fatal(err)
	}
	if len(pics) != 1 {
		t.Fatalf("want 1 picture, got %d", len(pics))
	}
}

func TestProtectSheetApplied(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)
	if err := w.ProtectSheet("Worksheet",
		`{"sheet":true,"password":"s3cret","formatCells":true,"selectLockedCells":false}`); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "protected.xlsx")
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, path, "<sheetProtection")
}

func TestAddChartSaves(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 5)
	spec := `{"type":"col","title":"Totals",
		"series":[{"name":"Worksheet!$A$1","categories":"Worksheet!$A$2:$A$5","values":"Worksheet!$B$2:$B$5"}],
		"legend":{"position":"bottom"}}`
	if err := w.AddChart("Worksheet", "E2", spec); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "chart.xlsx")
	if err := w.SaveXlsx(path); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "B3"); v != "3" {
		t.Errorf("data corrupted by chart: B3=%q", v)
	}
}

func TestCalculatedReadRows(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{
		mustCells(t, int64(2), int64(3)),
		{{Kind: compat.Formula, Str: "A1+B1"}},
	}); err != nil {
		t.Fatal(err)
	}
	out, _, err := w.ReadRows("Worksheet", 1, 2, false, true)
	if err != nil {
		t.Fatal(err)
	}
	if len(out) != 2 || out[1][0] != "5" {
		t.Errorf("calculated read = %+v, want row2 col1 = 5", out)
	}
}

func TestValidationBadSpecFailsEagerly(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.SetDataValidation("Worksheet", "A1", `{"type":"sparkles"}`); err == nil {
		t.Error("unknown validation type must fail at call time")
	}
	if err := w.SetConditionalFormat("Worksheet", "A1", `[{"type":"nope"}]`); err == nil {
		t.Error("unknown conditional type must fail at call time")
	}
	if err := w.AddChart("Worksheet", "A1", `{"type":"hexbin","series":[{}]}`); err == nil {
		t.Error("unknown chart type must fail at call time")
	}
}
