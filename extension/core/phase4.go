package core

import (
	"encoding/json"
	"fmt"
	"time"

	"github.com/xuri/excelize/v2"
)

// Phase 4.1 (PLAN.md §13): document properties, unmerge, merge introspection.
// Workbook encryption lives in the save/open paths (workbook.go), value
// binders and print titles are PHP-side.

// docPropsSpec mirrors PhpSpreadsheet's Document\Properties; company/manager
// live in app.xml.
type docPropsSpec struct {
	Title          string `json:"title"`
	Subject        string `json:"subject"`
	Creator        string `json:"creator"`
	LastModifiedBy string `json:"lastModifiedBy"`
	Description    string `json:"description"`
	Keywords       string `json:"keywords"`
	Category       string `json:"category"`
	Company        string `json:"company"` // app.xml; excelize exposes no Manager field
	Created        string `json:"created"` // RFC 3339
	Modified       string `json:"modified"`
}

// SetDocProps writes core (and, when set, app) document properties; the
// docProps parts are workbook-level, safe in any mode.
func (w *Workbook) SetDocProps(jsonSpec string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	var spec docPropsSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid document properties: %w", err)
	}
	if err := w.f.SetDocProps(&excelize.DocProperties{
		Title:          spec.Title,
		Subject:        spec.Subject,
		Creator:        spec.Creator,
		LastModifiedBy: spec.LastModifiedBy,
		Description:    spec.Description,
		Keywords:       spec.Keywords,
		Category:       spec.Category,
		Created:        spec.Created,
		Modified:       spec.Modified,
	}); err != nil {
		return err
	}
	if spec.Company != "" {
		return w.f.SetAppProps(&excelize.AppProperties{
			Application: "easy-excel",
			Company:     spec.Company,
		})
	}
	return nil
}

// customPropSpec is one custom document property. Type is PhpSpreadsheet's
// single-letter code (b/i/f/d/s); Remove deletes the property.
type customPropSpec struct {
	Name   string `json:"name"`
	Value  any    `json:"value"`
	Type   string `json:"type"`
	Remove bool   `json:"remove"`
}

// SetCustomProp writes (or removes) one custom document property.
func (w *Workbook) SetCustomProp(jsonSpec string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	var spec customPropSpec
	if err := json.Unmarshal([]byte(jsonSpec), &spec); err != nil {
		return fmt.Errorf("easy-excel: invalid custom property: %w", err)
	}
	if spec.Name == "" {
		return fmt.Errorf("easy-excel: custom property needs a name")
	}
	prop := excelize.CustomProperty{Name: spec.Name}
	if !spec.Remove {
		prop.Value = customPropValue(spec.Value, spec.Type)
	}
	return w.f.SetCustomProps(prop)
}

// customPropValue coerces a JSON value to the Go type excelize stores per the
// PhpSpreadsheet type code (JSON numbers arrive as float64).
func customPropValue(v any, typ string) any {
	switch typ {
	case "b":
		b, _ := v.(bool)
		return b
	case "i":
		if f, ok := v.(float64); ok {
			return int32(f) // excelize accepts int32, not int
		}
	case "f":
		if f, ok := v.(float64); ok {
			return f
		}
	case "d":
		if s, ok := v.(string); ok {
			if t, err := time.Parse(time.RFC3339, s); err == nil {
				return t
			}
		}
	}
	if s, ok := v.(string); ok {
		return s
	}
	return fmt.Sprintf("%v", v)
}

// UnmergeCells removes a merge. Merges queued for the StreamWriter but not
// yet applied are simply dropped; everything else unmerges at save.
func (w *Workbook) UnmergeCells(sheet, ref string) error {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return err
	}
	tl, br, err := splitRange(ref)
	if err != nil {
		return err
	}
	kept := st.preMerges[:0]
	dropped := false
	for _, m := range st.preMerges {
		if m[0] == tl && m[1] == br {
			dropped = true
			continue
		}
		kept = append(kept, m)
	}
	st.preMerges = kept
	if dropped {
		return nil
	}
	if st.random() {
		if err := w.mutable(); err != nil {
			return err
		}
		return w.f.UnmergeCell(sheet, tl, br)
	}
	st.pending = append(st.pending, pendingOp{kind: opUnmerge, ref: tl + ":" + br})
	return nil
}

// Merges returns the sheet's merged ranges ("A1:C3"). Reading degrades a
// streaming sheet, like every other read.
func (w *Workbook) Merges(sheet string) ([]string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return nil, errClosed
	}
	st, err := w.state(sheet)
	if err != nil {
		return nil, err
	}
	if err := w.ensureRandom(sheet, st); err != nil {
		return nil, err
	}
	merged, err := w.f.GetMergeCells(sheet)
	if err != nil {
		return nil, err
	}
	out := make([]string, 0, len(merged))
	for _, m := range merged {
		out = append(out, m.GetStartAxis()+":"+m.GetEndAxis())
	}
	return out, nil
}
