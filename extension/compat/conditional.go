package compat

import (
	"encoding/json"
	"fmt"

	"github.com/xuri/excelize/v2"
)

// ConditionalRule is one translated conditional-formatting rule; Style (when
// non-nil) must be registered with NewConditionalStyle and its ID placed in
// Options.Format by the caller (style IDs are per-excelize.File).
type ConditionalRule struct {
	Options excelize.ConditionalFormatOptions
	Style   *excelize.Style
}

type conditionalSpec struct {
	Type       string     `json:"type"`     // cellIs | containsText | expression | colorScale | dataBar
	Operator   string     `json:"operator"` // PhpSpreadsheet Conditional::OPERATOR_*
	Conditions []any      `json:"conditions"`
	Style      StyleSpec  `json:"style"`
	ColorScale *colorSpec `json:"colorScale"`
	DataBar    *barSpec   `json:"dataBar"`
	StopIfTrue bool       `json:"stopIfTrue"`
}

type colorSpec struct {
	MinColor string `json:"minColor"`
	MidColor string `json:"midColor"`
	MaxColor string `json:"maxColor"`
}

type barSpec struct {
	Color string `json:"color"`
}

// PhpSpreadsheet cellIs operators → excelize criteria strings.
var conditionalCriteria = map[string]string{
	"equal":              "==",
	"notEqual":           "!=",
	"greaterThan":        ">",
	"greaterThanOrEqual": ">=",
	"lessThan":           "<",
	"lessThanOrEqual":    "<=",
	"between":            "between",
	"notBetween":         "not between",
	"containsText":       "containing",
	"notContains":        "not containing",
	"beginsWith":         "begins with",
	"endsWith":           "ends with",
}

// TranslateConditionals decodes the shim's JSON rule list (sent by
// Style::setConditionalStyles) into excelize options + pending styles.
func TranslateConditionals(jsonRules string) ([]ConditionalRule, error) {
	var specs []conditionalSpec
	if err := json.Unmarshal([]byte(jsonRules), &specs); err != nil {
		return nil, fmt.Errorf("easy-excel: invalid conditional spec: %w", err)
	}
	rules := make([]ConditionalRule, 0, len(specs))
	for i, s := range specs {
		rule := ConditionalRule{}
		rule.Options.StopIfTrue = s.StopIfTrue
		switch s.Type {
		case "cellIs", "containsText":
			rule.Options.Type = "cell"
			if s.Type == "containsText" {
				rule.Options.Type = "text"
			}
			crit, ok := conditionalCriteria[s.Operator]
			if !ok {
				return nil, fmt.Errorf("easy-excel: conditional rule %d: unsupported operator %q", i, s.Operator)
			}
			rule.Options.Criteria = crit
			if len(s.Conditions) > 0 {
				rule.Options.Value = specString(s.Conditions[0])
			}
			if len(s.Conditions) > 1 {
				// excelize encodes between-bounds as "min,max"
				rule.Options.Value += "," + specString(s.Conditions[1])
			}
		case "expression":
			rule.Options.Type = "formula"
			if len(s.Conditions) == 0 {
				return nil, fmt.Errorf("easy-excel: conditional rule %d: expression needs a condition", i)
			}
			rule.Options.Criteria = specString(s.Conditions[0])
		case "colorScale":
			if s.ColorScale == nil {
				return nil, fmt.Errorf("easy-excel: conditional rule %d: missing colorScale", i)
			}
			rule.Options.Criteria = "=" // excelize requires it for scales
			rule.Options.MinType, rule.Options.MaxType = "min", "max"
			rule.Options.MinColor = stripAlpha(s.ColorScale.MinColor)
			rule.Options.MaxColor = stripAlpha(s.ColorScale.MaxColor)
			rule.Options.Type = "2_color_scale"
			if s.ColorScale.MidColor != "" {
				rule.Options.Type = "3_color_scale"
				rule.Options.MidType = "percentile"
				rule.Options.MidValue = "50"
				rule.Options.MidColor = stripAlpha(s.ColorScale.MidColor)
			}
		case "dataBar":
			if s.DataBar == nil {
				return nil, fmt.Errorf("easy-excel: conditional rule %d: missing dataBar", i)
			}
			rule.Options.Type = "data_bar"
			rule.Options.Criteria = "="
			rule.Options.MinType, rule.Options.MaxType = "min", "max"
			rule.Options.BarColor = stripAlpha(s.DataBar.Color)
		default:
			return nil, fmt.Errorf("easy-excel: unsupported conditional type %q", s.Type)
		}
		if len(s.Style) > 0 {
			style, err := TranslateStyle(s.Style)
			if err != nil {
				return nil, fmt.Errorf("easy-excel: conditional rule %d: %w", i, err)
			}
			rule.Style = style
		}
		rules = append(rules, rule)
	}
	return rules, nil
}
