package compat

import (
	"encoding/json"
	"fmt"
	"sort"
	"strings"

	"github.com/xuri/excelize/v2"
)

// StyleSpec is a decoded PhpSpreadsheet style array (the applyFromArray
// shape): font / fill / borders / alignment / numberFormat / protection.
// The PHP shim sends it as JSON; merging two specs is a recursive map merge,
// which mirrors how PhpSpreadsheet layers partial styles onto a cell.
type StyleSpec map[string]any

// ParseStyleSpec decodes the JSON form sent over the ABI.
func ParseStyleSpec(jsonSpec string) (StyleSpec, error) {
	var spec StyleSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return nil, fmt.Errorf("easy-excel: invalid style spec: %w", err)
	}
	return spec, nil
}

// MergeSpec layers patch over base without mutating either; nested maps merge
// recursively, scalars from patch win.
func MergeSpec(base, patch StyleSpec) StyleSpec {
	out := make(StyleSpec, len(base)+len(patch))
	for k, v := range base {
		out[k] = v
	}
	for k, pv := range patch {
		bm, bok := out[k].(map[string]any)
		pm, pok := pv.(map[string]any)
		if bok && pok {
			out[k] = map[string]any(MergeSpec(bm, pm))
		} else {
			out[k] = pv
		}
	}
	return out
}

// CanonicalKey returns a deterministic string form of the spec, used as the
// style-interner cache key (encoding/json sorts map keys).
func (s StyleSpec) CanonicalKey() string {
	b, err := json.Marshal(s)
	if err != nil {
		return fmt.Sprintf("%v", map[string]any(s))
	}
	return string(b)
}

// PhpSpreadsheet border style constants → excelize border style indexes.
var borderStyles = map[string]int{
	"none":             0,
	"thin":             1,
	"medium":           2,
	"dashed":           3,
	"dotted":           4,
	"thick":            5,
	"double":           6,
	"hair":             7,
	"mediumDashed":     8,
	"dashDot":          9,
	"mediumDashDot":    10,
	"dashDotDot":       11,
	"mediumDashDotDot": 12,
	"slantDashDot":     13,
}

// PhpSpreadsheet fill pattern constants → excelize pattern indexes.
var fillPatterns = map[string]int{
	"none":            0,
	"solid":           1,
	"mediumGray":      2,
	"darkGray":        3,
	"lightGray":       4,
	"darkHorizontal":  5,
	"darkVertical":    6,
	"darkDown":        7,
	"darkUp":          8,
	"darkGrid":        9,
	"darkTrellis":     10,
	"lightHorizontal": 11,
	"lightVertical":   12,
	"lightDown":       13,
	"lightUp":         14,
	"lightGrid":       15,
	"lightTrellis":    16,
	"gray125":         17,
	"gray0625":        18,
}

// TranslateStyle converts a spec into an excelize.Style. Unknown components
// or values fail loudly (compatibility policy: never silently produce a
// different file).
func TranslateStyle(spec StyleSpec) (*excelize.Style, error) {
	style := &excelize.Style{}
	for key, raw := range spec {
		section, ok := raw.(map[string]any)
		if !ok {
			return nil, fmt.Errorf("easy-excel: style component %q must be an array", key)
		}
		var err error
		switch key {
		case "font":
			err = translateFont(section, style)
		case "fill":
			err = translateFill(section, style)
		case "borders":
			err = translateBorders(section, style)
		case "alignment":
			err = translateAlignment(section, style)
		case "numberFormat":
			err = translateNumberFormat(section, style)
		case "protection":
			err = translateProtection(section, style)
		case "quotePrefix": // tolerated no-op section from applyFromArray
		default:
			err = fmt.Errorf("easy-excel: unsupported style component %q (see COMPAT.md)", key)
		}
		if err != nil {
			return nil, err
		}
	}
	return style, nil
}

func translateFont(m map[string]any, style *excelize.Style) error {
	f := &excelize.Font{}
	for k, v := range m {
		switch k {
		case "bold":
			f.Bold = truthy(v)
		case "italic":
			f.Italic = truthy(v)
		case "strikethrough", "strike":
			f.Strike = truthy(v)
		case "name":
			f.Family = specString(v)
		case "size":
			n, ok := v.(float64)
			if !ok {
				return fmt.Errorf("easy-excel: font size must be a number")
			}
			f.Size = n
		case "underline":
			u := specString(v)
			switch u {
			case "none", "false", "":
				f.Underline = ""
			case "single", "double":
				f.Underline = u
			case "singleAccounting", "doubleAccounting":
				f.Underline = strings.TrimSuffix(u, "Accounting")
			default:
				return fmt.Errorf("easy-excel: unsupported font underline %q", u)
			}
		case "color":
			f.Color = colorOf(v)
		case "superscript":
			if truthy(v) {
				f.VertAlign = "superscript"
			}
		case "subscript":
			if truthy(v) {
				f.VertAlign = "subscript"
			}
		default:
			return fmt.Errorf("easy-excel: unsupported font property %q", k)
		}
	}
	style.Font = f
	return nil
}

func translateFill(m map[string]any, style *excelize.Style) error {
	fill := excelize.Fill{Type: "pattern", Pattern: 1}
	var start, end string
	for k, v := range m {
		switch k {
		case "fillType":
			t := specString(v)
			if strings.HasPrefix(t, "linear") || strings.HasPrefix(t, "path") {
				return fmt.Errorf("easy-excel: gradient fills are not supported (see COMPAT.md)")
			}
			p, ok := fillPatterns[t]
			if !ok {
				return fmt.Errorf("easy-excel: unsupported fill type %q", t)
			}
			fill.Pattern = p
		case "startColor", "color":
			start = colorOf(v)
		case "endColor":
			end = colorOf(v)
		case "rotation":
			// gradient-only; ignored for pattern fills like PhpSpreadsheet
		default:
			return fmt.Errorf("easy-excel: unsupported fill property %q", k)
		}
	}
	if start != "" {
		fill.Color = []string{start}
	} else if end != "" {
		fill.Color = []string{end}
	}
	style.Fill = fill
	return nil
}

func translateBorders(m map[string]any, style *excelize.Style) error {
	sides := map[string]map[string]any{}
	for k, v := range m {
		side, ok := v.(map[string]any)
		if !ok {
			return fmt.Errorf("easy-excel: border %q must be an array", k)
		}
		switch k {
		case "allBorders", "outline":
			for _, s := range []string{"left", "right", "top", "bottom"} {
				sides[s] = side
			}
		case "left", "right", "top", "bottom":
			sides[k] = side
		case "diagonal", "vertical", "horizontal", "diagonalDirection":
			return fmt.Errorf("easy-excel: border %q is not supported (see COMPAT.md)", k)
		default:
			return fmt.Errorf("easy-excel: unsupported border %q", k)
		}
	}
	keys := make([]string, 0, len(sides))
	for k := range sides {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k := range keys {
		side := sides[k]
		b := excelize.Border{Type: k, Color: "000000"}
		if v, ok := side["borderStyle"]; ok {
			s, found := borderStyles[specString(v)]
			if !found {
				return fmt.Errorf("easy-excel: unsupported border style %q", specString(v))
			}
			b.Style = s
		}
		if v, ok := side["color"]; ok {
			b.Color = colorOf(v)
		}
		style.Border = append(style.Border, b)
	}
	return nil
}

func translateAlignment(m map[string]any, style *excelize.Style) error {
	a := &excelize.Alignment{}
	for k, v := range m {
		switch k {
		case "horizontal":
			h := specString(v)
			if h == "general" {
				h = ""
			}
			a.Horizontal = h
		case "vertical":
			a.Vertical = specString(v)
		case "wrapText":
			a.WrapText = truthy(v)
		case "shrinkToFit":
			a.ShrinkToFit = truthy(v)
		case "textRotation":
			n, _ := v.(float64)
			a.TextRotation = int(n)
		case "indent":
			n, _ := v.(float64)
			a.Indent = int(n)
		case "readOrder":
			n, _ := v.(float64)
			a.ReadingOrder = uint64(n)
		default:
			return fmt.Errorf("easy-excel: unsupported alignment property %q", k)
		}
	}
	style.Alignment = a
	return nil
}

func translateNumberFormat(m map[string]any, style *excelize.Style) error {
	for k, v := range m {
		switch k {
		case "formatCode":
			code := specString(v)
			style.CustomNumFmt = &code
		default:
			return fmt.Errorf("easy-excel: unsupported numberFormat property %q", k)
		}
	}
	return nil
}

func translateProtection(m map[string]any, style *excelize.Style) error {
	p := &excelize.Protection{Locked: true} // OOXML default
	for k, v := range m {
		switch k {
		case "locked":
			p.Locked = protectionFlag(v, true)
		case "hidden":
			p.Hidden = protectionFlag(v, false)
		default:
			return fmt.Errorf("easy-excel: unsupported protection property %q", k)
		}
	}
	style.Protection = p
	return nil
}

// protectionFlag accepts PhpSpreadsheet's string constants ('protected',
// 'unprotected', 'inherit') as well as plain booleans.
func protectionFlag(v any, inherit bool) bool {
	switch t := v.(type) {
	case bool:
		return t
	case string:
		switch t {
		case "protected":
			return true
		case "unprotected":
			return false
		}
	}
	return inherit
}

// colorOf accepts 'FF0000', 'FFFF0000' (ARGB), or PhpSpreadsheet's
// ['rgb' => ...] / ['argb' => ...] array form.
func colorOf(v any) string {
	switch t := v.(type) {
	case string:
		return stripAlpha(t)
	case map[string]any:
		if rgb, ok := t["rgb"]; ok {
			return stripAlpha(specString(rgb))
		}
		if argb, ok := t["argb"]; ok {
			return stripAlpha(specString(argb))
		}
	}
	return ""
}

func stripAlpha(c string) string {
	c = strings.TrimPrefix(c, "#")
	if len(c) == 8 {
		return c[2:]
	}
	return c
}

func truthy(v any) bool {
	b, ok := v.(bool)
	return ok && b
}

func specString(v any) string {
	if s, ok := v.(string); ok {
		return s
	}
	return fmt.Sprintf("%v", v)
}
