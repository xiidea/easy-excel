package compat

import (
	"testing"
)

func TestParseAndTranslateFullSpec(t *testing.T) {
	spec, err := ParseStyleSpec(`{
		"font": {"bold": true, "italic": true, "size": 14, "name": "Arial",
		         "underline": "single", "strikethrough": false,
		         "color": {"argb": "FFFF0000"}},
		"fill": {"fillType": "solid", "startColor": {"rgb": "FFFF00"}},
		"borders": {"allBorders": {"borderStyle": "thin", "color": {"rgb": "333333"}}},
		"alignment": {"horizontal": "center", "vertical": "top", "wrapText": true,
		              "textRotation": 45, "indent": 2},
		"numberFormat": {"formatCode": "0.00%"},
		"protection": {"locked": "unprotected", "hidden": true}
	}`)
	if err != nil {
		t.Fatal(err)
	}
	style, err := TranslateStyle(spec)
	if err != nil {
		t.Fatal(err)
	}
	if !style.Font.Bold || !style.Font.Italic || style.Font.Size != 14 {
		t.Errorf("font flags wrong: %+v", style.Font)
	}
	if style.Font.Family != "Arial" || style.Font.Underline != "single" {
		t.Errorf("font name/underline wrong: %+v", style.Font)
	}
	if style.Font.Color != "FF0000" {
		t.Errorf("ARGB alpha not stripped: %q", style.Font.Color)
	}
	if style.Fill.Pattern != 1 || len(style.Fill.Color) != 1 || style.Fill.Color[0] != "FFFF00" {
		t.Errorf("fill wrong: %+v", style.Fill)
	}
	if len(style.Border) != 4 {
		t.Fatalf("allBorders should expand to 4 sides, got %d", len(style.Border))
	}
	for _, b := range style.Border {
		if b.Style != 1 || b.Color != "333333" {
			t.Errorf("border side wrong: %+v", b)
		}
	}
	if style.Alignment.Horizontal != "center" || !style.Alignment.WrapText ||
		style.Alignment.TextRotation != 45 || style.Alignment.Indent != 2 {
		t.Errorf("alignment wrong: %+v", style.Alignment)
	}
	if style.CustomNumFmt == nil || *style.CustomNumFmt != "0.00%" {
		t.Errorf("number format wrong: %v", style.CustomNumFmt)
	}
	if style.Protection.Locked || !style.Protection.Hidden {
		t.Errorf("protection wrong: %+v", style.Protection)
	}
}

func TestTranslateBorderStyles(t *testing.T) {
	for name, want := range map[string]int{"thin": 1, "medium": 2, "thick": 5, "double": 6, "slantDashDot": 13} {
		spec := StyleSpec{"borders": map[string]any{"top": map[string]any{"borderStyle": name}}}
		style, err := TranslateStyle(spec)
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		if len(style.Border) != 1 || style.Border[0].Style != want {
			t.Errorf("%s: got %+v want style %d", name, style.Border, want)
		}
	}
}

func TestTranslateRejectsUnknown(t *testing.T) {
	for _, spec := range []StyleSpec{
		{"sparkles": map[string]any{}},
		{"font": map[string]any{"glow": true}},
		{"fill": map[string]any{"fillType": "linear"}},
		{"borders": map[string]any{"diagonal": map[string]any{"borderStyle": "thin"}}},
	} {
		if _, err := TranslateStyle(spec); err == nil {
			t.Errorf("expected error for %v", spec)
		}
	}
}

func TestMergeSpecLayersNestedMaps(t *testing.T) {
	base := StyleSpec{
		"font":         map[string]any{"bold": true, "size": float64(10)},
		"numberFormat": map[string]any{"formatCode": "0.00"},
	}
	patch := StyleSpec{
		"font": map[string]any{"italic": true, "size": float64(12)},
		"fill": map[string]any{"fillType": "solid"},
	}
	merged := MergeSpec(base, patch)
	font := merged["font"].(map[string]any)
	if font["bold"] != true || font["italic"] != true || font["size"] != float64(12) {
		t.Errorf("font merge wrong: %v", font)
	}
	if merged["numberFormat"] == nil || merged["fill"] == nil {
		t.Errorf("sections lost: %v", merged)
	}
	// inputs untouched
	if _, ok := base["fill"]; ok {
		t.Error("base mutated")
	}
	if bf := base["font"].(map[string]any); bf["italic"] != nil {
		t.Error("base font mutated")
	}
}

func TestCanonicalKeyDeterministic(t *testing.T) {
	a := StyleSpec{"font": map[string]any{"bold": true, "size": float64(10)}}
	b := StyleSpec{"font": map[string]any{"size": float64(10), "bold": true}}
	if a.CanonicalKey() != b.CanonicalKey() {
		t.Error("key should not depend on map order")
	}
}

func TestColorForms(t *testing.T) {
	for _, tc := range []struct{ in, want string }{
		{`{"font": {"color": "FF0000"}}`, "FF0000"},
		{`{"font": {"color": "FFFF0000"}}`, "FF0000"},
		{`{"font": {"color": {"rgb": "00FF00"}}}`, "00FF00"},
		{`{"font": {"color": {"argb": "FF0000FF"}}}`, "0000FF"},
	} {
		spec, err := ParseStyleSpec(tc.in)
		if err != nil {
			t.Fatal(err)
		}
		style, err := TranslateStyle(spec)
		if err != nil {
			t.Fatal(err)
		}
		if style.Font.Color != tc.want {
			t.Errorf("%s: got %q want %q", tc.in, style.Font.Color, tc.want)
		}
	}
}
