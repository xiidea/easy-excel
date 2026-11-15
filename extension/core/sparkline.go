package core

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"reflect"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"

	"github.com/xiidea/easy-excel/extension/compat"
)

// Sparklines (PhpSpreadsheet 5.10 Worksheet\Sparkline). excelize writes them
// into the worksheet's extLst, which the StreamWriter does not carry, so they
// queue and apply on flush like the rest of the non-streamable surface
// (COMPAT.md divergence 11).

// AddSparklineGroup queues one sparkline group (JSON spec, see
// compat.TranslateSparklineGroup). The spec is validated now so a bad group
// fails at the call site rather than at save.
func (w *Workbook) AddSparklineGroup(sheet, jsonSpec string) error {
	if _, err := compat.TranslateSparklineGroup(jsonSpec); err != nil {
		return err
	}
	return w.queueOp(sheet, pendingOp{kind: opSparkline, s1: jsonSpec})
}

// applyOpSparkline executes a queued sparkline group in random-access mode.
func (w *Workbook) applyOpSparkline(sheet string, op pendingOp) error {
	if op.kind != opSparkline {
		return fmt.Errorf("easy-excel: unhandled op kind %d", op.kind)
	}
	opts, err := compat.TranslateSparklineGroup(op.s1)
	if err != nil {
		return err
	}
	return w.f.AddSparkline(sheet, opts)
}

// extURISparklineGroups identifies the worksheet extLst entry that carries
// x14:sparklineGroups.
const extURISparklineGroups = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"

// LoadedSparklineXML returns the raw sparkline <ext> element a loaded
// worksheet carried in its source file, or "" when it had none (and always
// for workbooks created in memory). The shim parses it with
// Reader\Xlsx\Sparklines so getSparklineGroupCollection() reports the
// file's groups.
func (w *Workbook) LoadedSparklineXML(sheet string) (string, error) {
	w.mu.Lock()
	defer w.mu.Unlock()
	if w.closed {
		return "", errClosed
	}
	if _, err := w.state(sheet); err != nil {
		return "", err
	}
	return w.loadedSparklines[sheet], nil
}

// scanLoadedSparklines collects each worksheet's sparkline <ext> element at
// open time. It is a byte search over the parts excelize already holds in
// memory, not an XML parse, so files without sparklines pay one scan per
// sheet. Parts above excelize's in-memory limit (UnzipXMLSizeLimit) live in
// temp files excelize does not expose; those sheets report no loaded
// sparklines, although the file's own sparklines are still preserved on
// save (COMPAT.md divergence 35).
func scanLoadedSparklines(f *excelize.File) map[string]string {
	out := map[string]string{}
	for sheet, part := range pkgSheetParts(f) {
		data := pkgPart(f, part)
		if ext := extElement(data, extURISparklineGroups); ext != "" {
			out[sheet] = declareNamespaces(ext, ancestorNamespaces(data, bytes.Index(data, []byte(ext))))
		}
	}
	return out
}

// startTagNamespaces returns the xmlns:prefix declarations on the first start
// tag in data, nil if there is none.
func startTagNamespaces(data []byte) []xml.Attr {
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.RawToken()
		if err != nil {
			return nil
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		var out []xml.Attr
		for _, a := range start.Attr {
			if a.Name.Space == "xmlns" {
				out = append(out, a)
			}
		}
		return out
	}
}

// ancestorNamespaces lists the prefixes in scope for the <ext> element that
// starts at extStart in a worksheet part, nearest declaration first: an ext
// sits in the worksheet's extLst, so those are the extLst's own declarations
// and then the root element's.
func ancestorNamespaces(part []byte, extStart int) []xml.Attr {
	var out []xml.Attr
	if extStart > 0 {
		if lt := bytes.LastIndexByte(part[:extStart], '<'); lt >= 0 {
			out = startTagNamespaces(part[lt:extStart])
		}
	}
	return append(out, startTagNamespaces(part)...)
}

// declareNamespaces makes an element cut out of a part self-contained: each
// prefix in inScope (x14, xr2, a prefixed main namespace, ...) is declared on
// the element unless the element, or an earlier entry, already declares it.
// Without them a parser sees every such prefix as undefined.
func declareNamespaces(element string, inScope []xml.Attr) string {
	nameEnd := strings.IndexAny(element, " \t\r\n/>")
	if nameEnd < 0 || len(inScope) == 0 {
		return element
	}
	declared := map[string]bool{}
	for _, a := range startTagNamespaces([]byte(element)) {
		declared[a.Name.Local] = true
	}
	var decls strings.Builder
	for _, a := range inScope {
		if declared[a.Name.Local] {
			continue
		}
		declared[a.Name.Local] = true
		decls.WriteString(" xmlns:" + a.Name.Local + `="`)
		_ = xml.EscapeText(&decls, []byte(a.Value))
		decls.WriteByte('"')
	}
	return element[:nameEnd] + decls.String() + element[nameEnd:]
}

// liveSparklineXML returns the sparkline <ext> element sheet's worksheet
// holds right now, "" when it has none. Unlike loadedSparklines this includes
// groups added since the file was opened. ok is false when the worksheet
// cannot be inspected.
func (w *Workbook) liveSparklineXML(sheet string) (ext string, ok bool) {
	for id, name := range w.f.GetSheetMap() {
		if name != sheet {
			continue
		}
		// a sheet made by NewSheet lives at this path; excelize keeps the
		// extLst as raw inner XML in an unexported type, hence the reflection
		ws, found := w.f.Sheet.Load("xl/worksheets/sheet" + strconv.Itoa(id) + ".xml")
		if !found {
			return "", false
		}
		v := reflect.Indirect(reflect.ValueOf(ws))
		if v.Kind() != reflect.Struct {
			return "", false
		}
		lst := v.FieldByName("ExtLst")
		if !lst.IsValid() || lst.Kind() != reflect.Pointer {
			return "", false
		}
		if lst.IsNil() {
			return "", true
		}
		raw := lst.Elem().FieldByName("Ext")
		if !raw.IsValid() || raw.Kind() != reflect.String {
			return "", false
		}
		return extElement([]byte(raw.String()), extURISparklineGroups), true
	}
	return "", false
}

// pkgSheetParts maps sheet names to worksheet part paths from the workbook
// and its rels as excelize loaded them (f.Pkg), so it is only accurate at
// open time, before any rename.
func pkgSheetParts(f *excelize.File) map[string]string {
	out := map[string]string{}
	wbRaw, relsRaw := pkgPart(f, "xl/workbook.xml"), pkgPart(f, "xl/_rels/workbook.xml.rels")
	if wbRaw == nil || relsRaw == nil {
		return out
	}
	var wb wbXML
	var rels relsXML
	if xml.Unmarshal(wbRaw, &wb) != nil || xml.Unmarshal(relsRaw, &rels) != nil {
		return out
	}
	targets := map[string]string{}
	for _, r := range rels.Relationship {
		targets[r.ID] = r.Target
	}
	for _, s := range wb.Sheets.Sheet {
		if target, ok := targets[s.RID]; ok {
			out[s.Name] = resolvePart("xl/workbook.xml", target)
		}
	}
	return out
}

// resolvePart resolves a relationship target against the part that owns the
// rels: relative targets are relative to that part's directory, absolute
// ones are package-rooted.
func resolvePart(owner, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	return path.Join(path.Dir(owner), target)
}

func pkgPart(f *excelize.File, name string) []byte {
	v, ok := f.Pkg.Load(name)
	if !ok {
		return nil
	}
	b, _ := v.([]byte)
	return b
}

// extElement cuts the <ext uri="…"> element (any prefix) out of a worksheet
// part. ext elements do not nest, so the first matching close tag after the
// opening one ends it.
func extElement(data []byte, uri string) string {
	at := bytes.Index(data, []byte(uri))
	if at < 0 {
		return ""
	}
	start := bytes.LastIndexByte(data[:at], '<')
	if start < 0 {
		return ""
	}
	nameEnd := start + 1
	for nameEnd < at && data[nameEnd] != ' ' && data[nameEnd] != '\t' && data[nameEnd] != '\n' && data[nameEnd] != '\r' {
		nameEnd++
	}
	qname := string(data[start+1 : nameEnd])
	if qname != "ext" && !strings.HasSuffix(qname, ":ext") {
		return ""
	}
	closeTag := []byte("</" + qname + ">")
	end := bytes.Index(data[at:], closeTag)
	if end < 0 {
		return ""
	}
	return string(data[start : at+end+len(closeTag)])
}
