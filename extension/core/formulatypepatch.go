package core

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"regexp"
)

// Numeric formula results must not be typed "str".
//
// excelize's SetCellFormula ends with an unconditional `c.T, c.IS = "str", nil`
// and exposes no way to reset it: every public value setter calls
// removeFormula(), so writing the value afterwards deletes the formula. A
// cached numeric result therefore lands as
//
//	<c r="C1" s="1" t="str"><f>SUM(A1:B1)</f><v>2500</v></c>
//
// In OOXML `t="str"` means "formula returning a string"; the numeric type is
// `n`, the default, written by omitting the attribute. The consequence is not
// cosmetic: excelize's own reader takes a `case "str": return c.V` path that
// bypasses formattedValue, so a #,##0 total renders 2500 rather than 2,500 for
// any consumer that trusts the cache.
//
// The saved container is patched the same way streaming auto-filters are
// (filterpatch.go). Crucially the patch is driven by the *exact cell
// references* cacheFormulaResults() wrote, never by scanning for
// numeric-looking values: a genuine string result can look numeric —
// `TEXT(A1,"0000")` caches "0042" — and retyping it would silently turn it
// into 42. Only cells this engine cached in this save are touched, so formulas
// already present in a loaded workbook are left exactly as they were.

// cellRefAttr extracts the r="…" reference from a <c …> start tag.
var cellRefAttr = regexp.MustCompile(`\sr="([A-Z]+[0-9]+)"`)

// dropStrType removes the t="str" attribute from a single <c …> start tag.
var dropStrType = regexp.MustCompile(`\s+t="str"`)

// patchNumericFormulaTypes rewrites the saved container, dropping t="str" from
// exactly the cells listed in cached (sheet part path → set of cell refs).
// Worksheet parts with nothing to change are copied compressed-as-is, so a
// workbook with one cached total does not pay a full recompress of every sheet.
func patchNumericFormulaTypes(src io.ReaderAt, size int64, dst io.Writer, cached map[string]map[string]bool) error {
	zr, err := zip.NewReader(src, size)
	if err != nil {
		return err
	}
	parts, err := sheetPartPaths(zr)
	if err != nil {
		return err
	}
	// Re-key by part path: cached is keyed by sheet name.
	byPart := make(map[string]map[string]bool, len(cached))
	for sheet, cells := range cached {
		if len(cells) == 0 {
			continue
		}
		part, ok := parts[sheet]
		if !ok {
			return fmt.Errorf("easy-excel: no worksheet part for sheet %q", sheet)
		}
		byPart[part] = cells
	}

	zw := zip.NewWriter(dst)
	for _, f := range zr.File {
		cells, patch := byPart[f.Name]
		if !patch {
			if err := copyRaw(zw, f); err != nil {
				return err
			}
			continue
		}
		r, err := f.Open()
		if err != nil {
			return err
		}
		b, err := io.ReadAll(r)
		r.Close()
		if err != nil {
			return err
		}
		w, err := zw.Create(f.Name)
		if err != nil {
			return err
		}
		if _, err := w.Write(retypeCells(b, cells)); err != nil {
			return err
		}
	}
	return zw.Close()
}

// copyRaw copies a zip entry without recompressing it.
func copyRaw(zw *zip.Writer, f *zip.File) error {
	raw, err := f.OpenRaw()
	if err != nil {
		return err
	}
	hdr := f.FileHeader
	w, err := zw.CreateRaw(&hdr)
	if err != nil {
		return err
	}
	_, err = io.Copy(w, raw)
	return err
}

// retypeCells drops t="str" from the <c> start tags whose r="…" is in cells.
//
// Walks the start tags rather than matching whole elements: a cell's body may
// contain anything (shared formulas are self-closing, inline strings nest
// elements), and the reference is all that decides whether to rewrite.
func retypeCells(b []byte, cells map[string]bool) []byte {
	var out bytes.Buffer
	out.Grow(len(b))
	rest := b
	for {
		i := bytes.Index(rest, []byte("<c "))
		if i < 0 {
			out.Write(rest)
			return out.Bytes()
		}
		j := bytes.IndexByte(rest[i:], '>')
		if j < 0 { // malformed; emit the remainder untouched
			out.Write(rest)
			return out.Bytes()
		}
		tagEnd := i + j + 1
		out.Write(rest[:i])
		tag := rest[i:tagEnd]
		if m := cellRefAttr.FindSubmatch(tag); m != nil && cells[string(m[1])] {
			tag = dropStrType.ReplaceAll(tag, nil)
		}
		out.Write(tag)
		rest = rest[tagEnd:]
	}
}
