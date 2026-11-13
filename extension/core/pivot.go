package core

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"path"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
)

// Pivot tables (PhpSpreadsheet 5.10 Worksheet\PivotTable). PhpSpreadsheet
// regenerates its in-memory pivot tables on every save, so this is a
// replace-the-set op rather than an append: the shim sends a sheet's full
// list of generated tables, and the tables written for that sheet by an
// earlier save are deleted first. Tables that came in with a loaded file are
// never touched. Resending an unchanged set is a no-op: excelize's
// DeletePivotTable unlinks a table but leaves its parts in the package, so
// replacing only when something changed keeps the common save-twice case
// clean (COMPAT.md divergence 34). Like the rest of the non-streamable surface the op queues
// and applies on flush (COMPAT.md divergence 11); excelize reads the source
// header row at apply time, which is after every stream has been flushed.

// pivotSet records what the last applied op wrote to a sheet.
type pivotSet struct {
	spec  string   // the op's JSON, to detect an unchanged resend
	names []string // tables to delete before the next replacement
}

// SetPivotTables queues the replacement of sheet's generated pivot tables
// with jsonSpecs (a JSON list, see compat.TranslatePivotTable). Specs are
// validated now so a bad table fails at the call site.
func (w *Workbook) SetPivotTables(sheet, jsonSpecs string) error {
	var specs []json.RawMessage
	if err := json.Unmarshal([]byte(jsonSpecs), &specs); err != nil {
		return fmt.Errorf("easy-excel: invalid pivot table list: %w", err)
	}
	seen := map[string]bool{}
	for _, raw := range specs {
		opts, err := compat.TranslatePivotTable(sheet, string(raw))
		if err != nil {
			return err
		}
		if seen[opts.Name] {
			return fmt.Errorf("easy-excel: duplicate pivot table name %q on %s", opts.Name, sheet)
		}
		seen[opts.Name] = true
	}
	return w.queueOp(sheet, pendingOp{kind: opPivotTables, s1: jsonSpecs})
}

// applyOpPivotTables executes a queued pivot-table replacement in
// random-access mode.
func (w *Workbook) applyOpPivotTables(sheet string, op pendingOp) error {
	if op.kind != opPivotTables {
		return fmt.Errorf("easy-excel: unhandled op kind %d", op.kind)
	}
	prev, written := w.pivots[sheet]
	if written && prev.spec == op.s1 {
		return nil
	}
	var specs []json.RawMessage
	if err := json.Unmarshal([]byte(op.s1), &specs); err != nil {
		return err
	}
	// live names the generated tables actually in the file as the op
	// progresses. A failure part-way records exactly those under a spec that
	// never matches, so the next op deletes them instead of duplicating them.
	live := append([]string(nil), prev.names...)
	record := func(spec string) {
		if len(live) == 0 {
			delete(w.pivots, sheet)
			return
		}
		w.pivots[sheet] = pivotSet{spec: spec, names: live}
	}
	for len(live) > 0 {
		if err := w.f.DeletePivotTable(sheet, live[0]); err != nil {
			record("")
			return err
		}
		live = live[1:]
	}
	// What is left on the sheet came with the loaded file. Deletion is by
	// name, so a generated table sharing one of those names would take the
	// file's own table with it on the next replacement. A sheet excelize
	// cannot list (non-worksheet cache sources) skips the check.
	foreign := map[string]bool{}
	if existing, err := w.f.GetPivotTables(sheet); err == nil {
		for _, pt := range existing {
			foreign[pt.Name] = true
		}
	}
	for _, raw := range specs {
		opts, err := compat.TranslatePivotTable(sheet, string(raw))
		if err == nil && foreign[opts.Name] {
			err = fmt.Errorf("easy-excel: pivot table %q already exists on %s", opts.Name, sheet)
		}
		if err == nil {
			if err = w.f.AddPivotTable(opts); err != nil {
				err = fmt.Errorf("easy-excel: pivot table %q: %w", opts.Name, err)
			}
		}
		if err != nil {
			record("")
			return err
		}
		live = append(live, opts.Name)
	}
	record(op.s1)
	return nil
}

const (
	relTypePivotTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
	relTypePivotCache = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
)

// loadedPivot is one pivot table a source file carried: the raw
// pivotTableDefinition part and its pivotCacheDefinition ("" if missing).
type loadedPivot struct {
	Table string `json:"table"`
	Cache string `json:"cache"`
}

// LoadedPivotTablesJSON returns the pivot tables a loaded worksheet carried
// in its source file as a JSON list of raw parts, "[]" when none (and always
// for workbooks created in memory). The shim parses them with
// Reader\Xlsx\PivotTableReader.
func (w *Workbook) LoadedPivotTablesJSON(sheet string) (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return "", errClosed
	}
	if _, err := w.state(sheet); err != nil {
		return "", err
	}
	out, err := json.Marshal(w.loadedPivots[sheet])
	if err != nil || string(out) == "null" {
		return "[]", err
	}
	return string(out), nil
}

// scanLoadedPivots collects each worksheet's pivot table and cache parts at
// open time by following the sheet and pivot-table rels in f.Pkg. Pivot parts
// are small, so they are always held in memory.
func scanLoadedPivots(f *excelize.File) map[string][]loadedPivot {
	out := map[string][]loadedPivot{}
	for sheet, part := range pkgSheetParts(f) {
		for _, table := range relTargets(f, part, relTypePivotTable) {
			raw := pkgPart(f, table)
			if raw == nil {
				continue
			}
			p := loadedPivot{Table: string(raw)}
			if caches := relTargets(f, table, relTypePivotCache); len(caches) > 0 {
				p.Cache = string(pkgPart(f, caches[0]))
			}
			out[sheet] = append(out[sheet], p)
		}
	}
	return out
}

// relTargets lists the parts owner relates to with relType, resolved to
// package paths.
func relTargets(f *excelize.File, owner, relType string) []string {
	raw := pkgPart(f, path.Join(path.Dir(owner), "_rels", path.Base(owner)+".rels"))
	if raw == nil {
		return nil
	}
	var rels relsXML
	if xml.Unmarshal(raw, &rels) != nil {
		return nil
	}
	var out []string
	for _, r := range rels.Relationship {
		if r.Type == relType {
			out = append(out, resolvePart(owner, r.Target))
		}
	}
	return out
}
