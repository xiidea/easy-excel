package compat

import (
	"encoding/json"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// richRun is one decoded rich-text run: text plus an optional font spec in
// the same shape TranslateStyle's "font" section accepts.
type richRun struct {
	Text string         `json:"text"`
	Font map[string]any `json:"font"`
}

// TranslateRichText decodes the shim's JSON run list into excelize rich-text
// runs, reusing the font translation so run formatting matches cell styling.
func TranslateRichText(jsonRuns string) ([]excelize.RichTextRun, error) {
	var runs []richRun
	if err := json.Unmarshal([]byte(jsonRuns), &runs); err != nil {
		return nil, fmt.Errorf("easy-excel: invalid rich text: %w", err)
	}
	if len(runs) == 0 {
		return nil, fmt.Errorf("easy-excel: rich text needs at least one run")
	}
	out := make([]excelize.RichTextRun, 0, len(runs))
	for _, r := range runs {
		run := excelize.RichTextRun{Text: r.Text}
		if len(r.Font) > 0 {
			style := &excelize.Style{}
			if err := translateFont(r.Font, style); err != nil {
				return nil, err
			}
			run.Font = style.Font
		}
		out = append(out, run)
	}
	return out, nil
}
