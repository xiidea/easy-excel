package core

import (
	"bytes"
	"encoding/base64"
	"image"
	"image/png"
	"path/filepath"
	"testing"

	"github.com/xiidea/easy-excel/extension/compat"
)

func TestRichTextRoundTrip(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)
	runs := `[{"text":"Total: ","font":{"bold":true}},
	          {"text":"42","font":{"color":{"rgb":"FF0000"},"italic":true}}]`
	if err := w.SetRichText("Worksheet", "A1", runs); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "rich.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	got, err := f.GetCellRichText("Worksheet", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if len(got) != 2 {
		t.Fatalf("want 2 runs, got %d", len(got))
	}
	if got[0].Text != "Total: " || got[0].Font == nil || !got[0].Font.Bold {
		t.Errorf("run 0 wrong: %+v", got[0])
	}
	if got[1].Text != "42" || got[1].Font == nil || !got[1].Font.Italic || got[1].Font.Color != "FF0000" {
		t.Errorf("run 1 wrong: %+v font=%+v", got[1], got[1].Font)
	}
	if v, _ := f.GetCellValue("Worksheet", "A1"); v != "Total: 42" {
		t.Errorf("plain text = %q, want Total: 42", v)
	}
}

func TestRichTextWinsOverStreamedValue(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	// stream a plain value into A1, then queue rich text for it
	fillRows(t, w, "Worksheet", 1, 5) // A1 = "a"
	if err := w.SetRichText("Worksheet", "A1", `[{"text":"rich"}]`); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "richwin.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "A1"); v != "rich" {
		t.Errorf("A1 = %q, want rich (rich text applies after the stream)", v)
	}
}

func TestAddImageBytes(t *testing.T) {
	img := image.NewRGBA(image.Rect(0, 0, 30, 15))
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatal(err)
	}
	data := base64.StdEncoding.EncodeToString(buf.Bytes())

	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)
	spec := `{"data":"` + data + `","extension":".png","name":"mem","width":60}`
	if err := w.AddImageBytes("Worksheet", "C2", spec); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "membytes.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	pics, err := f.GetPictures("Worksheet", "C2")
	if err != nil {
		t.Fatal(err)
	}
	if len(pics) != 1 {
		t.Fatalf("want 1 picture, got %d", len(pics))
	}
	if pics[0].Extension != ".png" {
		t.Errorf("extension = %q", pics[0].Extension)
	}
}

func TestAutoFilterWithColumnRules(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	if err := w.WriteRows("Worksheet", 1, 1, [][]compat.Cell{
		mustCells(t, "Region", "Sales"),
		mustCells(t, "East", int64(2500)),
		mustCells(t, "West", int64(1200)),
	}); err != nil {
		t.Fatal(err)
	}
	cols := `[{"column":"A","expression":"x == East"},{"column":"B","expression":"x > 2000"}]`
	if err := w.AutoFilterWithColumns("Worksheet", "A1:B3", cols); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "filtercols.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	// column rules require the model path, never the container patch
	if w.Degraded() {
		// acceptable: column-rule filter forced a degrade. The patch
		// would have produced no FilterColumn entries.
	}
	assertSheetXMLContains(t, path, `<filterColumn colId="0">`)
	assertSheetXMLContains(t, path, `<customFilters>`)
}

func TestAutoFilterColumnsBadExpression(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 3)
	if err := w.SaveXlsx(filepath.Join(t.TempDir(), "x.xlsx"), ""); err != nil {
		t.Fatal(err)
	}
	// loaded-into-model path surfaces excelize's expression validation
	err = w.AutoFilterWithColumns("Worksheet", "A1:C3", `[{"column":"A","expression":"garbage tokens here too many"}]`)
	if err == nil {
		t.Error("invalid filter expression should error")
	}
}
