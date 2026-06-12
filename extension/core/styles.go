package core

import (
	"sync"

	"github.com/xuri/excelize/v2"

	"github.com/ronisaha/easy-excel/extension/compat"
)

// styleInterner memoizes excelize style IDs per canonical style spec
// (PLAN.md §6): repeated applyFromArray-style calls hit a map instead of
// growing the stylesheet. Reset on degrade because style IDs belong to the
// underlying excelize.File instance.
type styleInterner struct {
	mu     sync.Mutex
	bySpec map[string]int
}

// specID returns the excelize style ID for a spec, creating and caching it on
// first use. The cache key is the spec's canonical JSON form.
func (si *styleInterner) specID(f *excelize.File, spec compat.StyleSpec) (int, error) {
	key := spec.CanonicalKey()
	si.mu.Lock()
	defer si.mu.Unlock()
	if id, ok := si.bySpec[key]; ok {
		return id, nil
	}
	style, err := compat.TranslateStyle(spec)
	if err != nil {
		return 0, err
	}
	id, err := f.NewStyle(style)
	if err != nil {
		return 0, err
	}
	if si.bySpec == nil {
		si.bySpec = make(map[string]int)
	}
	si.bySpec[key] = id
	return id, nil
}

func (si *styleInterner) numFmtID(f *excelize.File, code string) (int, error) {
	return si.specID(f, compat.StyleSpec{
		"numberFormat": map[string]any{"formatCode": code},
	})
}

func (si *styleInterner) reset() {
	si.mu.Lock()
	si.bySpec = nil
	si.mu.Unlock()
}
