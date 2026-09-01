package compat

import (
	"testing"
)

const baseSeries = `"series":[{"name":"S","categories":"Sheet1!$A$1:$A$3","values":"Sheet1!$B$1:$B$3"}]`

func TestTranslateChartAxisMinMax(t *testing.T) {
	c, err := TranslateChart(`{"type":"col",` + baseSeries + `,"yAxis":{"minimum":0,"maximum":100}}`)
	if err != nil {
		t.Fatal(err)
	}
	if c.YAxis.Minimum == nil || *c.YAxis.Minimum != 0 {
		t.Errorf("YAxis.Minimum = %v, want 0", c.YAxis.Minimum)
	}
	if c.YAxis.Maximum == nil || *c.YAxis.Maximum != 100 {
		t.Errorf("YAxis.Maximum = %v, want 100", c.YAxis.Maximum)
	}
}

// A zero minimum must survive: it is meaningfully different from "unset",
// which is why the spec uses pointers.
func TestTranslateChartAxisZeroIsNotUnset(t *testing.T) {
	c, err := TranslateChart(`{"type":"col",` + baseSeries + `,"yAxis":{"minimum":0}}`)
	if err != nil {
		t.Fatal(err)
	}
	if c.YAxis.Minimum == nil {
		t.Fatal("an explicit minimum of 0 must not read as unset")
	}
	if *c.YAxis.Minimum != 0 {
		t.Errorf("Minimum = %v, want 0", *c.YAxis.Minimum)
	}

	c2, err := TranslateChart(`{"type":"col",` + baseSeries + `,"yAxis":{}}`)
	if err != nil {
		t.Fatal(err)
	}
	if c2.YAxis.Minimum != nil {
		t.Errorf("an omitted minimum must stay nil, got %v", *c2.YAxis.Minimum)
	}
}

func TestTranslateChartAxisLabelsNone(t *testing.T) {
	c, err := TranslateChart(`{"type":"col",` + baseSeries + `,"xAxis":{"labels":"none"}}`)
	if err != nil {
		t.Fatal(err)
	}
	if !c.XAxis.None {
		t.Error(`labels "none" must suppress the axis`)
	}
}

// "low"/"high"/"nextTo" have no excelize equivalent: accepted, not an error,
// and must not accidentally suppress the axis.
func TestTranslateChartAxisLabelsAcceptedPositions(t *testing.T) {
	for _, pos := range []string{"low", "high", "nextTo", ""} {
		c, err := TranslateChart(`{"type":"col",` + baseSeries + `,"xAxis":{"labels":"` + pos + `"}}`)
		if err != nil {
			t.Fatalf("labels %q: %v", pos, err)
		}
		if c.XAxis.None {
			t.Errorf("labels %q must not suppress the axis", pos)
		}
	}
}

func TestTranslateChartAxisRejectsUnknownLabels(t *testing.T) {
	if _, err := TranslateChart(`{"type":"col",` + baseSeries + `,"xAxis":{"labels":"sideways"}}`); err == nil {
		t.Error("an unknown label position must fail at call time")
	}
}

func TestTranslateChartGridlines(t *testing.T) {
	c, err := TranslateChart(`{"type":"col",` + baseSeries +
		`,"yAxis":{"majorGridlines":true,"minorGridlines":true}}`)
	if err != nil {
		t.Fatal(err)
	}
	if !c.YAxis.MajorGridLines {
		t.Error("major gridlines not set")
	}
	if !c.YAxis.MinorGridLines {
		t.Error("minor gridlines not set")
	}
}

func TestTranslateChartAxisExtras(t *testing.T) {
	c, err := TranslateChart(`{"type":"col",` + baseSeries +
		`,"yAxis":{"majorUnit":25,"logBase":10,"reverseOrder":true,"numFmt":"0.00%","fontColor":"FF0000"}}`)
	if err != nil {
		t.Fatal(err)
	}
	if c.YAxis.MajorUnit != 25 {
		t.Errorf("MajorUnit = %v, want 25", c.YAxis.MajorUnit)
	}
	if c.YAxis.LogBase != 10 {
		t.Errorf("LogBase = %v, want 10", c.YAxis.LogBase)
	}
	if !c.YAxis.ReverseOrder {
		t.Error("ReverseOrder not set")
	}
	if c.YAxis.NumFmt.CustomNumFmt != "0.00%" {
		t.Errorf("NumFmt = %q, want 0.00%%", c.YAxis.NumFmt.CustomNumFmt)
	}
	if c.YAxis.Font.Color != "FF0000" {
		t.Errorf("Font.Color = %q, want FF0000", c.YAxis.Font.Color)
	}
}

func TestTranslateChartShowValues(t *testing.T) {
	on, err := TranslateChart(`{"type":"col",` + baseSeries + `,"showValues":true}`)
	if err != nil {
		t.Fatal(err)
	}
	if !on.PlotArea.ShowVal {
		t.Error("showValues true must set PlotArea.ShowVal")
	}

	off, err := TranslateChart(`{"type":"col",` + baseSeries + `}`)
	if err != nil {
		t.Fatal(err)
	}
	if off.PlotArea.ShowVal {
		t.Error("showValues defaults to false")
	}
}

// An axis block must not disturb what already worked before wave 5.4.
func TestTranslateChartAxisDoesNotBreakBasics(t *testing.T) {
	c, err := TranslateChart(`{"type":"col","title":"T",` + baseSeries +
		`,"legend":{"position":"bottom"},"xAxisTitle":"X","yAxisTitle":"Y"` +
		`,"yAxis":{"majorGridlines":true},"width":480,"height":290}`)
	if err != nil {
		t.Fatal(err)
	}
	if len(c.Series) != 1 {
		t.Fatalf("series = %d, want 1", len(c.Series))
	}
	if c.Legend.Position != "bottom" {
		t.Errorf("legend = %q", c.Legend.Position)
	}
	if len(c.Title.Paragraph) == 0 || c.Title.Paragraph[0].Text != "T" {
		t.Error("title lost")
	}
	if len(c.XAxis.Title.Paragraph) == 0 || c.XAxis.Title.Paragraph[0].Text != "X" {
		t.Error("x axis title lost when an axis block is present")
	}
	if len(c.YAxis.Title.Paragraph) == 0 || c.YAxis.Title.Paragraph[0].Text != "Y" {
		t.Error("y axis title lost when an axis block is present")
	}
	if c.Dimension.Width != 480 || c.Dimension.Height != 290 {
		t.Error("dimensions lost")
	}
}

func TestTranslateChartNoAxisBlockIsInert(t *testing.T) {
	c, err := TranslateChart(`{"type":"col",` + baseSeries + `}`)
	if err != nil {
		t.Fatal(err)
	}
	if c.XAxis.None || c.YAxis.None {
		t.Error("no axis block must leave axes visible")
	}
	if c.YAxis.MajorGridLines || c.YAxis.MinorGridLines {
		t.Error("no axis block must leave gridlines off")
	}
	if c.YAxis.Minimum != nil || c.YAxis.Maximum != nil {
		t.Error("no axis block must leave bounds unset")
	}
}
