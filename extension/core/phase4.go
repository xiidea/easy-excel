package core

import (
	"encoding/json"
	"fmt"

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
