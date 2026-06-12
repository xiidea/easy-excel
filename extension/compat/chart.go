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
	Width  uint `json:"width"`
	Height uint `json:"height"`
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
		chart.Title = []excelize.RichTextRun{{Text: spec.Title}}
	}
	switch spec.Legend.Position {
	case "":
	case "none", "top", "bottom", "left", "right":
		chart.Legend.Position = spec.Legend.Position
	default:
		return nil, fmt.Errorf("easy-excel: unsupported legend position %q", spec.Legend.Position)
	}
	if spec.Width > 0 {
		chart.Dimension.Width = spec.Width
	}
	if spec.Height > 0 {
		chart.Dimension.Height = spec.Height
	}
	return chart, nil
}
