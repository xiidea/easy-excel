package core

import (
	"archive/zip"
	"io"
	"path/filepath"
	"strings"
	"testing"
)

func TestChartTitlesSurviveExcelizeUpgrade(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "titles.xlsx")
	fillRows(t, w, "Worksheet", 1, 3)
	spec := `{"type":"col","title":"Revenue by region",` +
		`"series":[{"name":"S","categories":"Worksheet!$A$1:$A$3","values":"Worksheet!$B$1:$B$3"}],` +
		`"xAxisTitle":"Region","yAxisTitle":"Amount"}`
	if err := w.AddChart("Worksheet", "D1", spec); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.OpenReader(path)
	if err != nil {
		t.Fatal(err)
	}
	defer zr.Close()
	var xml string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/charts/chart") {
			rc, _ := f.Open()
			b, _ := io.ReadAll(rc)
			rc.Close()
			xml = string(b)
		}
	}
	if xml == "" {
		t.Fatal("no chart part")
	}
	for _, want := range []string{"Revenue by region", "Region", "Amount"} {
		if !strings.Contains(xml, want) {
			t.Errorf("chart XML lost title text %q after the v2.11 ChartTitle migration", want)
		}
	}
}
