package core

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

// Streaming auto-filter (Phase 3). excelize's StreamWriter cannot emit
// <autoFilter>, so Phase 2 degraded at save — measured at ~10× the styled 1M
// write. Instead, the saved container is patched: every entry is raw-copied
// (no recompression) except the target worksheets, whose XML gets
// `<autoFilter ref="…"/>` injected right after </sheetData> — the
// schema-valid position given what the StreamWriter emits (mergeCells and
// friends come later in the CT_Worksheet sequence).

// filterPatch maps a sheet name to its auto-filter range.
type filterPatch struct {
	sheet string
	ref   string
}

// workbook.xml / rels subsets, just enough to map sheet name → part path.
type wbXML struct {
	Sheets struct {
		Sheet []struct {
			Name string `xml:"name,attr"`
			RID  string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
		} `xml:"sheet"`
	} `xml:"sheets"`
}

type relsXML struct {
	Relationship []struct {
		ID     string `xml:"Id,attr"`
		Target string `xml:"Target,attr"`
	} `xml:"Relationship"`
}

func sheetPartPaths(zr *zip.Reader) (map[string]string, error) {
	readEntry := func(name string) ([]byte, error) {
		for _, f := range zr.File {
			if f.Name == name {
				r, err := f.Open()
				if err != nil {
					return nil, err
				}
				defer r.Close()
				return io.ReadAll(r)
			}
		}
		return nil, fmt.Errorf("easy-excel: %s missing from container", name)
	}
	wbRaw, err := readEntry("xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	relsRaw, err := readEntry("xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}
	var wb wbXML
	if err := xml.Unmarshal(wbRaw, &wb); err != nil {
		return nil, err
	}
	var rels relsXML
	if err := xml.Unmarshal(relsRaw, &rels); err != nil {
		return nil, err
	}
	byRID := make(map[string]string, len(rels.Relationship))
	for _, r := range rels.Relationship {
		byRID[r.ID] = r.Target
	}
	out := make(map[string]string, len(wb.Sheets.Sheet))
	for _, s := range wb.Sheets.Sheet {
		target, ok := byRID[s.RID]
		if !ok {
			continue
		}
		target = strings.TrimPrefix(target, "/xl/")
		target = strings.TrimPrefix(target, "./")
		out[s.Name] = "xl/" + strings.TrimPrefix(target, "xl/")
	}
	return out, nil
}

// patchAutoFilters copies the xlsx container from src to dst, injecting
// auto-filter elements into the listed sheets.
func patchAutoFilters(src io.ReaderAt, size int64, dst io.Writer, patches []filterPatch) error {
	zr, err := zip.NewReader(src, size)
	if err != nil {
		return err
	}
	parts, err := sheetPartPaths(zr)
	if err != nil {
		return err
	}
	inject := make(map[string]string, len(patches)) // part path → ref
	for _, p := range patches {
		part, ok := parts[p.sheet]
		if !ok {
			return fmt.Errorf("easy-excel: no worksheet part for sheet %q", p.sheet)
		}
		inject[part] = p.ref
	}
	zw := zip.NewWriter(dst)
	for _, f := range zr.File {
		ref, patch := inject[f.Name]
		if !patch {
			raw, err := f.OpenRaw()
			if err != nil {
				return err
			}
			hdr := f.FileHeader
			w, err := zw.CreateRaw(&hdr)
			if err != nil {
				return err
			}
			if _, err := io.Copy(w, raw); err != nil {
				return err
			}
			continue
		}
		r, err := f.Open()
		if err != nil {
			return err
		}
		w, err := zw.Create(f.Name)
		if err != nil {
			r.Close()
			return err
		}
		elem := `<autoFilter ref="` + xmlEscapeAttr(ref) + `"/>`
		if err := injectAfterSheetData(r, w, elem); err != nil {
			r.Close()
			return fmt.Errorf("easy-excel: patch %s: %w", f.Name, err)
		}
		r.Close()
	}
	return zw.Close()
}

func xmlEscapeAttr(s string) string {
	var b bytes.Buffer
	_ = xml.EscapeText(&b, []byte(s))
	return b.String()
}

// injectAfterSheetData streams src to dst inserting elem right after the
// first </sheetData> (or self-closing <sheetData/>), with chunk carry-over so
// the needle is found across read boundaries at constant memory.
func injectAfterSheetData(src io.Reader, dst io.Writer, elem string) error {
	needles := []string{"</sheetData>", "<sheetData/>"}
	const chunkSize = 256 << 10
	carryMax := len(needles[0]) - 1
	buf := make([]byte, chunkSize)
	var carry []byte
	injected := false
	for {
		n, readErr := src.Read(buf)
		if n > 0 {
			window := append(carry, buf[:n]...)
			if !injected {
				for _, needle := range needles {
					if i := bytes.Index(window, []byte(needle)); i >= 0 {
						at := i + len(needle)
						if _, err := dst.Write(window[:at]); err != nil {
							return err
						}
						if _, err := io.WriteString(dst, elem); err != nil {
							return err
						}
						window = window[at:]
						injected = true
						break
					}
				}
			}
			keep := 0
			if !injected && len(window) > carryMax {
				keep = carryMax
			} else if !injected {
				keep = len(window)
			}
			flushTo := len(window) - keep
			if _, err := dst.Write(window[:flushTo]); err != nil {
				return err
			}
			carry = append([]byte(nil), window[flushTo:]...)
		}
		if readErr == io.EOF {
			break
		}
		if readErr != nil {
			return readErr
		}
	}
	if len(carry) > 0 {
		if _, err := dst.Write(carry); err != nil {
			return err
		}
	}
	if !injected {
		return fmt.Errorf("sheetData element not found")
	}
	return nil
}
