package compat

import (
	"encoding/json"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// chartSpec is the JSON form of easy-excel's native chart API (PLAN.md §5:
// PhpSpreadsheet's chart object model is out of scope; this maps a compact
// declarative spec onto excelize.AddChart).
type chartSpec struct {
	Type   string `json:"type"`
	Title  string `json:"title"`
	Series []struct {
		Name       string `json:"name"`
		Categories string `json:"categories"`
		Values     string `json:"values"`
	} `json:"series"`
	Legend struct {
		Position string `json:"position"` // top | bottom | left | right | none
	} `json:"legend"`
	XAxisTitle string    `json:"xAxisTitle"`
	YAxisTitle string    `json:"yAxisTitle"`
	XAxis      *axisSpec `json:"xAxis"`
	YAxis      *axisSpec `json:"yAxis"`
	ShowValues bool      `json:"showValues"`
	Width      uint      `json:"width"`
	Height     uint      `json:"height"`
}

// axisSpec is the wave-5.4 mapping of PhpSpreadsheet's Chart\Axis and
// Chart\GridLines onto excelize.ChartAxis. Pointer fields distinguish "not
// set" from a deliberate zero, matching PhpSpreadsheet's nullable options.
type axisSpec struct {
	// "low" | "high" | "nextTo" | "none" — PhpSpreadsheet's AXIS_LABELS_*.
	// Only "none" has a direct excelize equivalent (ChartAxis.None); the
	// other placements are not modelled by excelize and are ignored.
	Labels        string   `json:"labels"`
	Minimum       *float64 `json:"minimum"`
	Maximum       *float64 `json:"maximum"`
	MajorUnit     *float64 `json:"majorUnit"`
	LogBase       *float64 `json:"logBase"`
	ReverseOrder  bool     `json:"reverseOrder"`
	MajorGridines bool     `json:"majorGridlines"`
	MinorGridines bool     `json:"minorGridlines"`
	NumFmt        string   `json:"numFmt"`
	FontColor     string   `json:"fontColor"`
}

// applyAxis folds an axisSpec onto an excelize.ChartAxis, leaving the title
// (set separately from xAxisTitle/yAxisTitle) untouched.
func applyAxis(dst *excelize.ChartAxis, spec *axisSpec) error {
	if spec == nil {
		return nil
	}
	switch spec.Labels {
	case "", "low", "high", "nextTo":
		// excelize has no tick-label-position field; only full suppression is
		// expressible, so these placements are accepted and ignored.
	case "none":
		dst.None = true
	default:
		return fmt.Errorf("easy-excel: unsupported axis label position %q", spec.Labels)
	}
	if spec.Minimum != nil {
		dst.Minimum = spec.Minimum
	}
	if spec.Maximum != nil {
		dst.Maximum = spec.Maximum
	}
	if spec.MajorUnit != nil {
		dst.MajorUnit = *spec.MajorUnit
	}
	if spec.LogBase != nil {
		dst.LogBase = *spec.LogBase
	}
	dst.ReverseOrder = spec.ReverseOrder
	dst.MajorGridLines = spec.MajorGridines
	dst.MinorGridLines = spec.MinorGridines
	if spec.NumFmt != "" {
		dst.NumFmt = excelize.ChartNumFmt{CustomNumFmt: spec.NumFmt}
	}
	if spec.FontColor != "" {
		dst.Font.Color = spec.FontColor
	}
	return nil
}

var chartTypes = map[string]excelize.ChartType{
	"area":       excelize.Area,
	"bar":        excelize.Bar,
	"barStacked": excelize.BarStacked,
	"col":        excelize.Col,
	"colStacked": excelize.ColStacked,
	"doughnut":   excelize.Doughnut,
	"line":       excelize.Line,
	"pie":        excelize.Pie,
	"radar":      excelize.Radar,
	"scatter":    excelize.Scatter,
}

// TranslateChart builds an excelize chart from the JSON spec.
func TranslateChart(jsonSpec string) (*excelize.Chart, error) {
	var spec chartSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return nil, fmt.Errorf("easy-excel: invalid chart spec: %w", err)
	}
	t, ok := chartTypes[spec.Type]
	if !ok {
		return nil, fmt.Errorf("easy-excel: unsupported chart type %q", spec.Type)
	}
	if len(spec.Series) == 0 {
		return nil, fmt.Errorf("easy-excel: chart needs at least one series")
	}
	chart := &excelize.Chart{Type: t}
	for _, s := range spec.Series {
		chart.Series = append(chart.Series, excelize.ChartSeries{
			Name:       s.Name,
			Categories: s.Categories,
			Values:     s.Values,
		})
	}
	if spec.Title != "" {
		chart.Title = excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: spec.Title}}}
	}
	switch spec.Legend.Position {
	case "":
	case "none", "top", "bottom", "left", "right":
		chart.Legend.Position = spec.Legend.Position
	default:
		return nil, fmt.Errorf("easy-excel: unsupported legend position %q", spec.Legend.Position)
	}
	if spec.XAxisTitle != "" {
		chart.XAxis.Title = excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: spec.XAxisTitle}}}
	}
	if spec.YAxisTitle != "" {
		chart.YAxis.Title = excelize.ChartTitle{Paragraph: []excelize.RichTextRun{{Text: spec.YAxisTitle}}}
	}
	if err := applyAxis(&chart.XAxis, spec.XAxis); err != nil {
		return nil, err
	}
	if err := applyAxis(&chart.YAxis, spec.YAxis); err != nil {
		return nil, err
	}
	// PhpSpreadsheet's Layout::setShowVal maps to excelize's plot-area flag;
	// the rest of Layout (manual plot-area geometry) has no excelize model.
	chart.PlotArea.ShowVal = spec.ShowValues
	if spec.Width > 0 {
		chart.Dimension.Width = spec.Width
	}
	if spec.Height > 0 {
		chart.Dimension.Height = spec.Height
	}
	return chart, nil
}
