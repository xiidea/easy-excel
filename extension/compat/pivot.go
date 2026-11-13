package compat

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

// pivotTableSpec is the JSON form of a generated PhpSpreadsheet 5.10
// Worksheet\PivotTable\PivotTable sent by the shim. Field grouping is not
// carried: excelize pivot tables have none (COMPAT.md divergence 33).
type pivotTableSpec struct {
	Name        string           `json:"name"`
	SourceSheet string           `json:"sourceSheet"`
	SourceRange string           `json:"sourceRange"` // e.g. A1:C100, header row included
	Location    string           `json:"location"`    // top-left cell, or a range
	Rows        []pivotFieldSpec `json:"rows"`
	Columns     []pivotFieldSpec `json:"columns"`
	Filters     []pivotFieldSpec `json:"filters"`
	Data        []pivotFieldSpec `json:"data"`
}

type pivotFieldSpec struct {
	Field    string `json:"field"`
	Subtotal string `json:"subtotal"` // data fields only: OOXML dataField@subtotal
	Caption  string `json:"caption"`  // data fields only
}

// OOXML dataField@subtotal values; excelize matches them case-insensitively
// but silently falls back to sum for anything else, so unknown names are
// rejected here instead.
var pivotSubtotals = map[string]bool{
	"average": true, "count": true, "countNums": true, "max": true, "min": true,
	"product": true, "stdDev": true, "stdDevp": true, "sum": true, "var": true, "varp": true,
}

// TranslatePivotTable maps a pivot table spec onto excelize options for a
// table placed on targetSheet.
func TranslatePivotTable(targetSheet, jsonSpec string) (*excelize.PivotTableOptions, error) {
	var spec pivotTableSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return nil, fmt.Errorf("easy-excel: invalid pivot table spec: %w", err)
	}
	if spec.Name == "" {
		return nil, fmt.Errorf("easy-excel: pivot table has no name")
	}
	if len(spec.Data) == 0 {
		return nil, fmt.Errorf("easy-excel: pivot table %q requires at least one data (value) field", spec.Name)
	}
	// excelize splits "Sheet!A1:B2" on '!' and takes the sheet name verbatim
	for _, sheet := range []string{spec.SourceSheet, targetSheet} {
		if sheet == "" || strings.Contains(sheet, "!") {
			return nil, fmt.Errorf("easy-excel: pivot table %q: unsupported sheet name %q", spec.Name, sheet)
		}
	}
	source := strings.ReplaceAll(spec.SourceRange, "$", "")
	tl, br, err := splitPivotRange(source)
	if err != nil {
		return nil, fmt.Errorf("easy-excel: pivot table %q source range: %w", spec.Name, err)
	}
	if tl == br {
		// excelize rejects it too, but only when the table is applied
		return nil, fmt.Errorf("easy-excel: pivot table %q source range %q is a single cell", spec.Name, spec.SourceRange)
	}
	location, err := pivotLocation(spec.Location)
	if err != nil {
		return nil, fmt.Errorf("easy-excel: pivot table %q location: %w", spec.Name, err)
	}

	// Defaults mirror what PhpSpreadsheet's writer emits: grand totals and
	// drill-down on (attributes omitted = true), PivotStyleLight16 with row
	// and column headers and the last column shown.
	opts := &excelize.PivotTableOptions{
		DataRange:           spec.SourceSheet + "!" + source,
		PivotTableRange:     targetSheet + "!" + location,
		Name:                spec.Name,
		RowGrandTotals:      true,
		ColGrandTotals:      true,
		ShowDrill:           true,
		UseAutoFormatting:   true,
		ItemPrintTitles:     true,
		ShowRowHeaders:      true,
		ShowColHeaders:      true,
		ShowLastColumn:      true,
		PivotTableStyleName: "PivotStyleLight16",
	}
	axis := func(fields []pivotFieldSpec) []excelize.PivotTableField {
		out := make([]excelize.PivotTableField, 0, len(fields))
		for _, f := range fields {
			out = append(out, excelize.PivotTableField{Data: f.Field, DefaultSubtotal: true})
		}
		return out
	}
	opts.Rows = axis(spec.Rows)
	opts.Columns = axis(spec.Columns)
	opts.Filter = axis(spec.Filters)
	for _, f := range spec.Data {
		subtotal := f.Subtotal
		if subtotal == "" {
			subtotal = "sum"
		}
		if !pivotSubtotals[subtotal] {
			return nil, fmt.Errorf("easy-excel: pivot table %q: unsupported subtotal %q for %q", spec.Name, f.Subtotal, f.Field)
		}
		opts.Data = append(opts.Data, excelize.PivotTableField{Data: f.Field, Subtotal: subtotal, Name: f.Caption})
	}
	return opts, nil
}

// pivotLocation turns PhpSpreadsheet's location into a range excelize
// accepts. The builder stores a degenerate "A3:A3" (the extent is recomputed
// on refresh), but excelize rejects single-cell ranges, so a lone cell is
// widened to the smallest two-by-two block anchored at it.
func pivotLocation(location string) (string, error) {
	location = strings.ReplaceAll(location, "$", "")
	tl, br, err := splitPivotRange(location)
	if err != nil {
		return "", err
	}
	if tl != br {
		return tl + ":" + br, nil
	}
	col, row, err := excelize.CellNameToCoordinates(tl)
	if err != nil {
		return "", err
	}
	end, err := excelize.CoordinatesToCellName(col+1, row+1)
	if err != nil {
		return "", err
	}
	return tl + ":" + end, nil
}

// splitPivotRange accepts "A1" or "A1:C9" and returns both corners.
func splitPivotRange(ref string) (string, string, error) {
	tl, br, found := strings.Cut(ref, ":")
	if !found {
		br = tl
	}
	for _, cell := range []string{tl, br} {
		if _, _, err := excelize.CellNameToCoordinates(cell); err != nil {
			return "", "", err
		}
	}
	return strings.ToUpper(tl), strings.ToUpper(br), nil
}
