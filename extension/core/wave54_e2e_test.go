package core

import (
	"archive/zip"
	"io"
	"path/filepath"
	"strings"
	"testing"
)

// End-to-end: the exact spec the shim builds for budget-service's
// ProgramVarianceReportManager must produce a real chart part.
func TestWave54ChartAxisEndToEnd(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "axis.xlsx")

	fillRows(t, w, "Worksheet", 1, 5)
	spec := `{"type":"bar","series":[{"name":"Sheet1!$A$1","categories":"Worksheet!$A$1:$A$5","values":"Worksheet!$B$1:$B$5"}],` +
		`"xAxis":{"labels":"none"},"yAxis":{"majorGridlines":true,"minimum":0,"maximum":100},` +
		`"showValues":true,"width":192,"height":280}`
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
	var chartXML string
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "xl/charts/chart") {
			rc, _ := f.Open()
			b, _ := io.ReadAll(rc)
			rc.Close()
			chartXML = string(b)
		}
	}
	if chartXML == "" {
		t.Fatal("no chart part written")
	}
	// excelize writes the chart part with unprefixed element names.
	for _, want := range []string{
		`<max val="100">`,   // yAxis maximum
		`<min val="0">`,     // yAxis minimum — an explicit zero
		"<majorGridlines>",  // yAxis gridlines
		`<showVal val="1">`, // layout showVal
	} {
		if !strings.Contains(chartXML, want) {
			t.Errorf("chart XML missing %s", want)
		}
	}
	// xAxis labels:"none" suppresses the category axis.
	if !strings.Contains(chartXML, `<delete val="1">`) {
		t.Error(`xAxis labels "none" must set delete=1 on the category axis`)
	}
}
