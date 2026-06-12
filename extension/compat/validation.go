package compat

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// validationSpec is the JSON form of PhpSpreadsheet's DataValidation sent by
// the shim (Cell::getDataValidation() / Worksheet::setDataValidation()).
type validationSpec struct {
	Type             string `json:"type"`
	Operator         string `json:"operator"`
	Formula1         string `json:"formula1"`
	Formula2         string `json:"formula2"`
	AllowBlank       bool   `json:"allowBlank"`
	ShowDropDown     bool   `json:"showDropDown"`
	ShowInputMessage bool   `json:"showInputMessage"`
	ShowErrorMessage bool   `json:"showErrorMessage"`
	ErrorStyle       string `json:"errorStyle"`
	ErrorTitle       string `json:"errorTitle"`
	Error            string `json:"error"`
	PromptTitle      string `json:"promptTitle"`
	Prompt           string `json:"prompt"`
}

// PhpSpreadsheet operator names → OOXML operator attribute (excelize uses the
// raw attribute values).
var validationOperators = map[string]string{
	"between":            "between",
	"notBetween":         "notBetween",
	"equal":              "equal",
	"notEqual":           "notEqual",
	"greaterThan":        "greaterThan",
	"greaterThanOrEqual": "greaterThanOrEqual",
	"lessThan":           "lessThan",
	"lessThanOrEqual":    "lessThanOrEqual",
}

var validationTypes = map[string]string{
	"list":       "list",
	"whole":      "whole",
	"decimal":    "decimal",
	"date":       "date",
	"time":       "time",
	"textLength": "textLength",
	"custom":     "custom",
}

// TranslateValidation builds an excelize DataValidation for a range from the
// shim's JSON spec.
func TranslateValidation(ref, jsonSpec string) (*excelize.DataValidation, error) {
	var spec validationSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return nil, fmt.Errorf("easy-excel: invalid validation spec: %w", err)
	}
	if spec.Type == "" || spec.Type == "none" {
		return nil, nil // cleared validation: nothing to add
	}
	t, ok := validationTypes[spec.Type]
	if !ok {
		return nil, fmt.Errorf("easy-excel: unsupported validation type %q", spec.Type)
	}
	dv := excelize.NewDataValidation(spec.AllowBlank)
	dv.Sqref = ref
	dv.Type = t
	if spec.Operator != "" {
		op, ok := validationOperators[spec.Operator]
		if !ok {
			return nil, fmt.Errorf("easy-excel: unsupported validation operator %q", spec.Operator)
		}
		dv.Operator = op
	}
	dv.Formula1 = validationFormula(t, spec.Formula1)
	dv.Formula2 = validationFormula(t, spec.Formula2)
	// PhpSpreadsheet's showDropDown=true means SUPPRESS the in-cell dropdown
	// (it maps to the OOXML showDropDown attribute), matching excelize
	dv.ShowDropDown = spec.ShowDropDown
	dv.ShowInputMessage = spec.ShowInputMessage
	dv.ShowErrorMessage = spec.ShowErrorMessage
	if spec.ErrorStyle != "" {
		s := spec.ErrorStyle // stop | warning | information
		dv.ErrorStyle = &s
	}
	setOpt := func(dst **string, v string) {
		if v != "" {
			s := v
			*dst = &s
		}
	}
	setOpt(&dv.ErrorTitle, spec.ErrorTitle)
	setOpt(&dv.Error, spec.Error)
	setOpt(&dv.PromptTitle, spec.PromptTitle)
	setOpt(&dv.Prompt, spec.Prompt)
	return dv, nil
}

// validationFormula normalizes PhpSpreadsheet formula syntax: list literals
// arrive as `"a,b,c"` (already quoted) or bare; range refs and other types
// pass through with any leading '=' stripped.
func validationFormula(typ, f string) string {
	f = strings.TrimPrefix(f, "=")
	if typ == "list" && f != "" && !strings.HasPrefix(f, `"`) && !isRangeRef(f) {
		return `"` + strings.ReplaceAll(f, `"`, `""`) + `"`
	}
	return f
}

func isRangeRef(f string) bool {
	if strings.ContainsAny(f, "!$:") {
		return true
	}
	_, _, err := excelize.CellNameToCoordinates(f)
	return err == nil
}
