package core

import (
	"path/filepath"
	"testing"
)

func TestCustomDocProps(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	fillRows(t, w, "Worksheet", 1, 2)
	for _, spec := range []string{
		`{"name":"Reviewed","value":true,"type":"b"}`,
		`{"name":"Revision","value":3,"type":"i"}`,
		`{"name":"Ratio","value":1.5,"type":"f"}`,
		`{"name":"Owner","value":"QA","type":"s"}`,
	} {
		if err := w.SetCustomProp(spec); err != nil {
			t.Fatalf("%s: %v", spec, err)
		}
	}
	if err := w.SetDocProps(`{"title":"T","created":"2026-06-13T10:00:00Z"}`); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(t.TempDir(), "props.xlsx")
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	f := reopen(t, path)
	props, err := f.GetCustomProps()
	if err != nil {
		t.Fatal(err)
	}
	got := map[string]any{}
	for _, p := range props {
		got[p.Name] = p.Value
	}
	if got["Reviewed"] != true || got["Owner"] != "QA" {
		t.Errorf("custom props wrong: %+v", got)
	}
	if _, ok := got["Revision"]; !ok {
		t.Errorf("Revision missing: %+v", got)
	}
	// removal
	if err := w.SetCustomProp(`{"name":"Owner","remove":true}`); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "props2.xlsx")
	if err := w.SaveXlsx(out, ""); err != nil {
		t.Fatal(err)
	}
	f2 := reopen(t, out)
	props2, _ := f2.GetCustomProps()
	for _, p := range props2 {
		if p.Name == "Owner" {
			t.Error("Owner should have been removed")
		}
	}
}
