// easy_excel.go is the FrankenPHP bridge: the flat, handle-first ABI exposed
// to PHP (PLAN.md §4). It contains no spreadsheet logic — it marshals
// arguments, dispatches to core, and encodes results.
//
// Error convention (the generator cannot throw PHP exceptions from Go):
//   - mutating functions return ?string — null on success, message on error;
//   - value-returning functions return array{0: mixed, 1: ?string};
//   - messages are prefixed "OVERLOADED:", "DENIED:" or "BADHANDLE:" so the
//     PHP shim can raise typed exceptions (see EasyExcel\Native::raise()).
//
// Build (Docker is the supported path, see Dockerfile / README):
//
//	GEN_STUB_SCRIPT=php-src/build/gen_stub.php frankenphp extension-init easy_excel.go
//	  # generates easy_excel.c/.h, _arginfo.h, _generated.go in-place
//	CGO_ENABLED=1 CGO_CFLAGS=$(php-config --includes) \
//	  CGO_LDFLAGS="$(php-config --ldflags) $(php-config --libs)" \
//	  xcaddy build --output frankenphp \
//	  --with github.com/dunglas/frankenphp/caddy \
//	  --with github.com/xiidea/easy-excel/extension=$PWD
//
// Note: the generator requires each Go parameter declared separately
// (no grouped `a, b int64`), and gofmt must not touch this file — it
// mangles //export_php: directives.
package easy_excel

// #include <Zend/zend_types.h>
import "C"

import (
	"errors"
	"fmt"
	"time"
	"unsafe"

	"github.com/dunglas/frankenphp"

	"github.com/xiidea/easy-excel/extension/compat"
	"github.com/xiidea/easy-excel/extension/core"
	"github.com/xiidea/easy-excel/extension/exio"
	"github.com/xiidea/easy-excel/extension/limits"
	"github.com/xiidea/easy-excel/extension/registry"
)

const version = "0.1.0"

// Process-wide state: one gate, one policy, one handle table
// (PLAN.md §7). The 10-minute idle TTL is the leak backstop for PHP code
// that never reaches the shim's shutdown handler.
var (
	gate    = limits.NewGate(limits.FromEnv())
	policy  = mustPolicy()
	env     = &core.Env{Gate: gate, Policy: policy}
	handles = registry.New(10 * time.Minute)
)

func mustPolicy() *exio.Policy {
	p, err := exio.FromEnv()
	if err != nil {
		panic(err)
	}
	return p
}

// --- result encoding ---------------------------------------------------------

func errString(err error) string {
	switch {
	case errors.Is(err, limits.ErrOverloaded):
		return "OVERLOADED: " + err.Error()
	case errors.Is(err, exio.ErrDenied):
		return "DENIED: " + err.Error()
	case errors.Is(err, registry.ErrNotFound):
		return "BADHANDLE: " + err.Error()
	default:
		return err.Error()
	}
}

// errOnly encodes the ?string convention for mutating functions.
func errOnly(err error) unsafe.Pointer {
	if err == nil {
		return nil
	}
	return frankenphp.PHPString(errString(err), false)
}

// pair encodes the array{0: mixed, 1: ?string} convention.
func pair(value any, err error) unsafe.Pointer {
	if err != nil {
		return frankenphp.PHPPackedArray([]any{nil, errString(err)})
	}
	return frankenphp.PHPPackedArray([]any{value, nil})
}

func workbook(handle int64) (*core.Workbook, error) {
	v, err := handles.Get(handle)
	if err != nil {
		return nil, err
	}
	wb, ok := v.(*core.Workbook)
	if !ok {
		return nil, registry.ErrNotFound
	}
	return wb, nil
}

func goStr(s *C.zend_string) string {
	return frankenphp.GoString(unsafe.Pointer(s))
}

// --- lifecycle ----------------------------------------------------------------

//export_php:function easy_excel_version(): string
func easy_excel_version() unsafe.Pointer {
	return frankenphp.PHPString(version, false)
}

//export_php:function easy_excel_new(): array
func easy_excel_new() unsafe.Pointer {
	wb, err := core.New(env)
	if err != nil {
		return pair(nil, err)
	}
	return pair(handles.Put(wb), nil)
}

//export_php:function easy_excel_open(string $path, string $password): array
func easy_excel_open(path *C.zend_string, password *C.zend_string) unsafe.Pointer {
	wb, err := core.Open(goStr(path), goStr(password), env)
	if err != nil {
		return pair(nil, err)
	}
	return pair(handles.Put(wb), nil)
}

//export_php:function easy_excel_close(int $handle): ?string
func easy_excel_close(handle int64) unsafe.Pointer {
	v, err := handles.Remove(handle)
	if err != nil {
		// closing a stale handle is a no-op, not an error: destructor +
		// shutdown handler + TTL janitor may race benignly
		return nil
	}
	if wb, ok := v.(*core.Workbook); ok {
		return errOnly(wb.Close())
	}
	return nil
}

// --- sheet management ----------------------------------------------------------

//export_php:function easy_excel_add_sheet(int $handle, string $name): array
func easy_excel_add_sheet(handle int64, name *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	idx, err := wb.AddSheet(goStr(name))
	return pair(int64(idx), err)
}

//export_php:function easy_excel_delete_sheet(int $handle, string $name): ?string
func easy_excel_delete_sheet(handle int64, name *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.DeleteSheet(goStr(name)))
}

//export_php:function easy_excel_rename_sheet(int $handle, string $old, string $new): ?string
func easy_excel_rename_sheet(handle int64, oldName *C.zend_string, newName *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.RenameSheet(goStr(oldName), goStr(newName)))
}

//export_php:function easy_excel_sheets(int $handle): array
func easy_excel_sheets(handle int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	names := wb.Sheets()
	list := make([]any, len(names))
	for i, n := range names {
		list[i] = n
	}
	return pair(list, nil)
}

//export_php:function easy_excel_set_active_sheet(int $handle, int $index): ?string
func easy_excel_set_active_sheet(handle int64, index int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetActiveSheet(int(index)))
}

//export_php:function easy_excel_active_sheet(int $handle): array
func easy_excel_active_sheet(handle int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	pos, name := wb.ActiveSheet()
	return pair([]any{int64(pos), name}, nil)
}

// --- write path -----------------------------------------------------------------

//export_php:function easy_excel_write_rows(int $handle, string $sheet, int $startRow, int $startCol, array $rows): ?string
func easy_excel_write_rows(handle int64, sheet *C.zend_string, startRow int64, startCol int64, rows *C.zend_array) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	rawRows, err := frankenphp.GoPackedArray[any](unsafe.Pointer(rows))
	if err != nil {
		return errOnly(fmt.Errorf("easy-excel: rows must be a packed array: %w", err))
	}
	batch := make([][]compat.Cell, len(rawRows))
	for i, rr := range rawRows {
		cells, ok := rr.([]any)
		if !ok {
			return errOnly(fmt.Errorf("easy-excel: row %d is not a packed array (got %T)", i, rr))
		}
		row := make([]compat.Cell, len(cells))
		for j, cv := range cells {
			c, err := compat.Decode(cv)
			if err != nil {
				return errOnly(fmt.Errorf("row %d col %d: %w", i, j, err))
			}
			row[j] = c
		}
		batch[i] = row
	}
	return errOnly(wb.WriteRows(goStr(sheet), int(startRow), int(startCol), batch))
}

//export_php:function easy_excel_set_cell(int $handle, string $sheet, string $cell, array $value): ?string
func easy_excel_set_cell(handle int64, sheet *C.zend_string, cell *C.zend_string, value *C.zend_array) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	// value is [scalar] for auto-binding or [marker, scalar] for explicit
	parts, err := frankenphp.GoPackedArray[any](unsafe.Pointer(value))
	if err != nil {
		return errOnly(err)
	}
	var c compat.Cell
	switch len(parts) {
	case 1:
		c, err = compat.Decode(parts[0])
	case 2:
		c, err = compat.Decode(parts)
	default:
		err = fmt.Errorf("easy-excel: cell value must be [value] or [marker, value]")
	}
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetCell(goStr(sheet), goStr(cell), c))
}

// --- read path --------------------------------------------------------------------

//export_php:function easy_excel_get_cell(int $handle, string $sheet, string $cell, int $mode): array
func easy_excel_get_cell(handle int64, sheet *C.zend_string, cell *C.zend_string, mode int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	v, err := wb.GetCell(goStr(sheet), goStr(cell), core.GetMode(mode))
	return pair(v, err)
}

//export_php:function easy_excel_read_rows(int $handle, string $sheet, int $startRow, int $maxRows, bool $raw, bool $calc): array
func easy_excel_read_rows(handle int64, sheet *C.zend_string, startRow int64, maxRows int64, raw bool, calc bool) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	rows, more, err := wb.ReadRows(goStr(sheet), int(startRow), int(maxRows), raw, calc)
	if err != nil {
		return pair(nil, err)
	}
	out := make([]any, len(rows))
	for i, r := range rows {
		cols := make([]any, len(r))
		for j, c := range r {
			cols[j] = c
		}
		out[i] = cols
	}
	return pair([]any{out, more}, nil)
}

//export_php:function easy_excel_dimensions(int $handle, string $sheet): array
func easy_excel_dimensions(handle int64, sheet *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	maxRow, maxCol, err := wb.Dimensions(goStr(sheet))
	return pair([]any{int64(maxRow), int64(maxCol)}, err)
}

// --- styling / structure -------------------------------------------------------------

//export_php:function easy_excel_set_number_format(int $handle, string $sheet, string $range, string $code): ?string
func easy_excel_set_number_format(handle int64, sheet *C.zend_string, ref *C.zend_string, code *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetNumberFormat(goStr(sheet), goStr(ref), goStr(code)))
}

//export_php:function easy_excel_merge_cells(int $handle, string $sheet, string $range): ?string
func easy_excel_merge_cells(handle int64, sheet *C.zend_string, ref *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.MergeCells(goStr(sheet), goStr(ref)))
}

//export_php:function easy_excel_apply_style(int $handle, string $sheet, string $range, string $styleJson): ?string
func easy_excel_apply_style(handle int64, sheet *C.zend_string, ref *C.zend_string, styleJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.ApplyStyle(goStr(sheet), goStr(ref), goStr(styleJson)))
}

//export_php:function easy_excel_set_col_width(int $handle, string $sheet, int $startCol, int $endCol, float $width): ?string
func easy_excel_set_col_width(handle int64, sheet *C.zend_string, startCol int64, endCol int64, width float64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetColWidth(goStr(sheet), int(startCol), int(endCol), width))
}

//export_php:function easy_excel_set_col_autosize(int $handle, string $sheet, int $startCol, int $endCol): ?string
func easy_excel_set_col_autosize(handle int64, sheet *C.zend_string, startCol int64, endCol int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetColAutoSize(goStr(sheet), int(startCol), int(endCol)))
}

//export_php:function easy_excel_set_row_height(int $handle, string $sheet, int $row, float $height): ?string
func easy_excel_set_row_height(handle int64, sheet *C.zend_string, row int64, height float64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetRowHeight(goStr(sheet), int(row), height))
}

//export_php:function easy_excel_freeze_panes(int $handle, string $sheet, string $topLeftCell): ?string
func easy_excel_freeze_panes(handle int64, sheet *C.zend_string, topLeftCell *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.FreezePanes(goStr(sheet), goStr(topLeftCell)))
}

//export_php:function easy_excel_auto_filter(int $handle, string $sheet, string $range): ?string
func easy_excel_auto_filter(handle int64, sheet *C.zend_string, ref *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.AutoFilter(goStr(sheet), goStr(ref)))
}

//export_php:function easy_excel_set_hyperlink(int $handle, string $sheet, string $cell, string $url, string $tooltip): ?string
func easy_excel_set_hyperlink(handle int64, sheet *C.zend_string, cell *C.zend_string, url *C.zend_string, tooltip *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetHyperlink(goStr(sheet), goStr(cell), goStr(url), goStr(tooltip)))
}

//export_php:function easy_excel_set_comment(int $handle, string $sheet, string $cell, string $author, string $text): ?string
func easy_excel_set_comment(handle int64, sheet *C.zend_string, cell *C.zend_string, author *C.zend_string, text *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetComment(goStr(sheet), goStr(cell), goStr(author), goStr(text)))
}

//export_php:function easy_excel_defined_name(int $handle, string $name, string $refersTo, string $scopeSheet): ?string
func easy_excel_defined_name(handle int64, name *C.zend_string, refersTo *C.zend_string, scopeSheet *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetDefinedName(goStr(name), goStr(refersTo), goStr(scopeSheet)))
}

//export_php:function easy_excel_page_setup(int $handle, string $sheet, string $orientation, int $paperSize, int $fitToWidth, int $fitToHeight): ?string
func easy_excel_page_setup(handle int64, sheet *C.zend_string, orientation *C.zend_string, paperSize int64, fitToWidth int64, fitToHeight int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetPageSetup(goStr(sheet), goStr(orientation), int(paperSize), int(fitToWidth), int(fitToHeight)))
}

//export_php:function easy_excel_set_validation(int $handle, string $sheet, string $range, string $validationJson): ?string
func easy_excel_set_validation(handle int64, sheet *C.zend_string, ref *C.zend_string, validationJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetDataValidation(goStr(sheet), goStr(ref), goStr(validationJson)))
}

//export_php:function easy_excel_set_conditional(int $handle, string $sheet, string $range, string $rulesJson): ?string
func easy_excel_set_conditional(handle int64, sheet *C.zend_string, ref *C.zend_string, rulesJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetConditionalFormat(goStr(sheet), goStr(ref), goStr(rulesJson)))
}

//export_php:function easy_excel_add_image(int $handle, string $sheet, string $cell, string $imageJson): ?string
func easy_excel_add_image(handle int64, sheet *C.zend_string, cell *C.zend_string, imageJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.AddImage(goStr(sheet), goStr(cell), goStr(imageJson)))
}

//export_php:function easy_excel_protect_sheet(int $handle, string $sheet, string $protectionJson): ?string
func easy_excel_protect_sheet(handle int64, sheet *C.zend_string, protectionJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.ProtectSheet(goStr(sheet), goStr(protectionJson)))
}

//export_php:function easy_excel_add_chart(int $handle, string $sheet, string $cell, string $chartJson): ?string
func easy_excel_add_chart(handle int64, sheet *C.zend_string, cell *C.zend_string, chartJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.AddChart(goStr(sheet), goStr(cell), goStr(chartJson)))
}

// --- save -------------------------------------------------------------------------

//export_php:function easy_excel_save_xlsx(int $handle, string $path, string $password): ?string
func easy_excel_save_xlsx(handle int64, path *C.zend_string, password *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SaveXlsx(goStr(path), goStr(password)))
}

//export_php:function easy_excel_doc_props(int $handle, string $propsJson): ?string
func easy_excel_doc_props(handle int64, propsJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetDocProps(goStr(propsJson)))
}

//export_php:function easy_excel_unmerge_cells(int $handle, string $sheet, string $range): ?string
func easy_excel_unmerge_cells(handle int64, sheet *C.zend_string, ref *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.UnmergeCells(goStr(sheet), goStr(ref)))
}

//export_php:function easy_excel_get_merges(int $handle, string $sheet): array
func easy_excel_get_merges(handle int64, sheet *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	refs, err := wb.Merges(goStr(sheet))
	if err != nil {
		return pair(nil, err)
	}
	list := make([]any, len(refs))
	for i, r := range refs {
		list[i] = r
	}
	return pair(list, nil)
}

//export_php:function easy_excel_get_style(int $handle, string $sheet, string $cell): array
func easy_excel_get_style(handle int64, sheet *C.zend_string, cell *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	spec, err := wb.GetStyleSpec(goStr(sheet), goStr(cell))
	return pair(spec, err)
}

//export_php:function easy_excel_get_validations(int $handle, string $sheet): array
func easy_excel_get_validations(handle int64, sheet *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	out, err := wb.ValidationsJSON(goStr(sheet))
	return pair(out, err)
}

//export_php:function easy_excel_get_conditionals(int $handle, string $sheet): array
func easy_excel_get_conditionals(handle int64, sheet *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	out, err := wb.ConditionalsJSON(goStr(sheet))
	return pair(out, err)
}

//export_php:function easy_excel_get_defined_names(int $handle): array
func easy_excel_get_defined_names(handle int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	out, err := wb.DefinedNamesJSON()
	return pair(out, err)
}

//export_php:function easy_excel_set_default_style(int $handle, string $styleJson): ?string
func easy_excel_set_default_style(handle int64, styleJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetDefaultStyle(goStr(styleJson)))
}

//export_php:function easy_excel_insert_rows(int $handle, string $sheet, int $row, int $count): ?string
func easy_excel_insert_rows(handle int64, sheet *C.zend_string, row int64, count int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.InsertRows(goStr(sheet), int(row), int(count)))
}

//export_php:function easy_excel_remove_rows(int $handle, string $sheet, int $row, int $count): ?string
func easy_excel_remove_rows(handle int64, sheet *C.zend_string, row int64, count int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.RemoveRows(goStr(sheet), int(row), int(count)))
}

//export_php:function easy_excel_insert_cols(int $handle, string $sheet, int $col, int $count): ?string
func easy_excel_insert_cols(handle int64, sheet *C.zend_string, col int64, count int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.InsertCols(goStr(sheet), int(col), int(count)))
}

//export_php:function easy_excel_remove_cols(int $handle, string $sheet, int $col, int $count): ?string
func easy_excel_remove_cols(handle int64, sheet *C.zend_string, col int64, count int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.RemoveCols(goStr(sheet), int(col), int(count)))
}

//export_php:function easy_excel_move_sheet(int $handle, string $sheet, int $index): ?string
func easy_excel_move_sheet(handle int64, sheet *C.zend_string, index int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.MoveSheetTo(goStr(sheet), int(index)))
}

//export_php:function easy_excel_copy_sheet(int $handle, string $from, string $newName): array
func easy_excel_copy_sheet(handle int64, from *C.zend_string, newName *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	idx, err := wb.CopySheetTo(goStr(from), goStr(newName))
	return pair(int64(idx), err)
}

//export_php:function easy_excel_sheet_view(int $handle, string $sheet, string $viewJson): ?string
func easy_excel_sheet_view(handle int64, sheet *C.zend_string, viewJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetSheetView(goStr(sheet), goStr(viewJson)))
}

//export_php:function easy_excel_header_footer(int $handle, string $sheet, string $hfJson): ?string
func easy_excel_header_footer(handle int64, sheet *C.zend_string, hfJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetHeaderFooter(goStr(sheet), goStr(hfJson)))
}

//export_php:function easy_excel_page_margins(int $handle, string $sheet, string $marginsJson): ?string
func easy_excel_page_margins(handle int64, sheet *C.zend_string, marginsJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetPageMargins(goStr(sheet), goStr(marginsJson)))
}

//export_php:function easy_excel_set_rich_text(int $handle, string $sheet, string $cell, string $runsJson): ?string
func easy_excel_set_rich_text(handle int64, sheet *C.zend_string, cell *C.zend_string, runsJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetRichText(goStr(sheet), goStr(cell), goStr(runsJson)))
}

//export_php:function easy_excel_add_image_bytes(int $handle, string $sheet, string $cell, string $imageJson): ?string
func easy_excel_add_image_bytes(handle int64, sheet *C.zend_string, cell *C.zend_string, imageJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.AddImageBytes(goStr(sheet), goStr(cell), goStr(imageJson)))
}

//export_php:function easy_excel_auto_filter_columns(int $handle, string $sheet, string $range, string $columnsJson): ?string
func easy_excel_auto_filter_columns(handle int64, sheet *C.zend_string, ref *C.zend_string, columnsJson *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.AutoFilterWithColumns(goStr(sheet), goStr(ref), goStr(columnsJson)))
}

//export_php:function easy_excel_save_csv(int $handle, string $path, string $sheet, string $delimiter, bool $crlf, bool $bom, bool $guardFormulas): ?string
func easy_excel_save_csv(handle int64, path *C.zend_string, sheet *C.zend_string, delimiter *C.zend_string, crlf bool, bom bool, guard bool) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	opts := core.CsvOptions{UseCRLF: crlf, UseBOM: bom, GuardFormula: guard}
	if d := goStr(delimiter); d != "" {
		opts.Delimiter = []rune(d)[0]
	}
	return errOnly(wb.SaveCsv(goStr(path), goStr(sheet), opts))
}

// --- observability ------------------------------------------------------------------

//export_php:function easy_excel_stats(): array
func easy_excel_stats() unsafe.Pointer {
	return frankenphp.PHPPackedArray([]any{
		int64(handles.Len()),
		gate.MemoryUsed(),
	})
}
