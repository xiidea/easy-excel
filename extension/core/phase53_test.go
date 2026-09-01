package core

import (
	"path/filepath"
	"testing"
)

func TestSetBreakRowWritesRowBreak(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "rowbreak.xlsx")

	fillRows(t, w, "Worksheet", 1, 30)
	if err := w.SetBreak("Worksheet", "A24", breakRow); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	// A row break at A24 splits above row 24 → rowBreaks entry id="23".
	assertSheetXMLContains(t, path, `<rowBreaks`)
	assertSheetXMLContains(t, path, `id="23"`)

	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "B30"); v != "30" {
		t.Errorf("B30 = %q, want 30 — data must survive the break", v)
	}
}

func TestSetBreakColumnWritesColBreak(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "colbreak.xlsx")

	fillRows(t, w, "Worksheet", 1, 5)
	if err := w.SetBreak("Worksheet", "C1", breakColumn); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, path, `<colBreaks`)
}

// A row break anchored at a non-A column must still be a pure row break:
// PhpSpreadsheet's setBreak('O24', BREAK_ROW) means "split above row 24",
// not "also split left of column O".
func TestSetBreakRowNormalisesColumn(t *testing.T) {
	for _, ref := range []string{"A24", "O24", "ZZ24"} {
		cell, err := breakCell(ref, breakRow)
		if err != nil {
			t.Fatal(err)
		}
		if cell != "A24" {
			t.Errorf("breakCell(%q, row) = %q, want A24", ref, cell)
		}
	}
	for _, ref := range []string{"O1", "O5", "O99"} {
		cell, err := breakCell(ref, breakColumn)
		if err != nil {
			t.Fatal(err)
		}
		if cell != "O1" {
			t.Errorf("breakCell(%q, col) = %q, want O1", ref, cell)
		}
	}
}

func TestSetBreakNoneRemoves(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "nobreak.xlsx")

	fillRows(t, w, "Worksheet", 1, 30)
	if err := w.SetBreak("Worksheet", "A24", breakRow); err != nil {
		t.Fatal(err)
	}
	if err := w.SetBreak("Worksheet", "A24", breakNone); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	// Removing the only break leaves no rowBreaks element with a live entry.
	f := reopen(t, path)
	if v, _ := f.GetCellValue("Worksheet", "B30"); v != "30" {
		t.Errorf("B30 = %q after break removal, want 30", v)
	}
}

func TestSetBreakRejectsBadInput(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()

	if err := w.SetBreak("Worksheet", "A1", 99); err == nil {
		t.Error("unknown break type must fail at call time")
	}
	if err := w.SetBreak("Worksheet", "not-a-cell", breakRow); err == nil {
		t.Error("invalid cell reference must fail at call time")
	}
}

func TestSetSelectionWritesSelection(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "selection.xlsx")

	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.SetSelection("Worksheet", "B2:D5"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, path, `sqref="B2:D5"`)
	assertSheetXMLContains(t, path, `activeCell="B2"`)
}

func TestSetSelectionSingleCell(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "selection1.xlsx")

	fillRows(t, w, "Worksheet", 1, 5)
	if err := w.SetSelection("Worksheet", "C3"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, path, `activeCell="C3"`)
}

// Selection must not silently undo a freeze: both ride the same pane record,
// so a naive SetPanes would drop the freeze set earlier.
func TestSetSelectionPreservesFreezePanes(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "freeze-and-select.xlsx")

	if err := w.FreezePanes("Worksheet", "A2"); err != nil {
		t.Fatal(err)
	}
	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.SetSelection("Worksheet", "B5"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}

	assertSheetXMLContains(t, path, `<pane`)
	assertSheetXMLContains(t, path, `ySplit="1"`)
	assertSheetXMLContains(t, path, `activeCell="B5"`)
}

func TestSetSelectionRejectsBadInput(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()

	if err := w.SetSelection("Worksheet", "nope!"); err == nil {
		t.Error("invalid selection must fail at call time")
	}
	if err := w.SetSelection("Worksheet", ""); err != nil {
		t.Errorf("empty selection is a no-op, got %v", err)
	}
}

func TestBreakAndSelectionOnUnknownSheet(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()

	if err := w.SetBreak("NoSuchSheet", "A1", breakRow); err == nil {
		t.Error("break on a missing sheet must fail")
	}
	if err := w.SetSelection("NoSuchSheet", "A1"); err == nil {
		t.Error("selection on a missing sheet must fail")
	}
}

// Review item 4(b): the mirror of TestSetSelectionPreservesFreezePanes —
// selection set *before* the freeze was discarded, because FreezePanes builds
// a fresh Panes with a nil Selection.
func TestFreezePanesPreservesSelection(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "select-then-freeze.xlsx")

	fillRows(t, w, "Worksheet", 1, 10)
	if err := w.SetSelection("Worksheet", "B5"); err != nil {
		t.Fatal(err)
	}
	if err := w.FreezePanes("Worksheet", "A2"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, path, `ySplit="1"`)
	assertSheetXMLContains(t, path, `activeCell="B5"`)
}

// Same ordering, but in random-access mode where FreezePanes calls SetPanes
// directly rather than queueing.
func TestFreezePanesPreservesSelectionRandomAccess(t *testing.T) {
	w, err := New(testEnv())
	if err != nil {
		t.Fatal(err)
	}
	defer w.Close()
	path := filepath.Join(t.TempDir(), "ra-select-then-freeze.xlsx")

	fillRows(t, w, "Worksheet", 1, 10)
	// force random access before the pane ops
	if _, _, err := w.ReadRows("Worksheet", 1, 1, false, false); err != nil {
		t.Fatal(err)
	}
	if err := w.SetSelection("Worksheet", "B5"); err != nil {
		t.Fatal(err)
	}
	if err := w.FreezePanes("Worksheet", "A2"); err != nil {
		t.Fatal(err)
	}
	if err := w.SaveXlsx(path, ""); err != nil {
		t.Fatal(err)
	}
	assertSheetXMLContains(t, path, `ySplit="1"`)
	assertSheetXMLContains(t, path, `activeCell="B5"`)
}
