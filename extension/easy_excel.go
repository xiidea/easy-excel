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
// Build (requires the frankenphp CLI and PHP ZTS dev headers, see README):
//
//	GEN_STUB_SCRIPT=php-src/build/gen_stub.php frankenphp extension-init easy_excel.go
//	CGO_ENABLED=1 CGO_CFLAGS=$(php-config --includes) \
//	  CGO_LDFLAGS="$(php-config --ldflags) $(php-config --libs)" \
//	  xcaddy build --output frankenphp --with github.com/ronisaha/easy-excel/extension/build
package easy_excel

// #include <Zend/zend_types.h>
import "C"

import (
	"errors"
	"fmt"
	"time"
	"unsafe"

	"github.com/dunglas/frankenphp"

	"github.com/ronisaha/easy-excel/extension/compat"
	"github.com/ronisaha/easy-excel/extension/core"
	"github.com/ronisaha/easy-excel/extension/exio"
	"github.com/ronisaha/easy-excel/extension/limits"
	"github.com/ronisaha/easy-excel/extension/registry"
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

//export_php:function easy_excel_open(string $path): array
func easy_excel_open(path *C.zend_string) unsafe.Pointer {
	wb, err := core.Open(goStr(path), env)
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
func easy_excel_rename_sheet(handle int64, oldName, newName *C.zend_string) unsafe.Pointer {
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
func easy_excel_set_active_sheet(handle, index int64) unsafe.Pointer {
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
func easy_excel_write_rows(handle int64, sheet *C.zend_string, startRow, startCol int64, rows *C.zend_array) unsafe.Pointer {
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
func easy_excel_set_cell(handle int64, sheet, cell *C.zend_string, value *C.zend_array) unsafe.Pointer {
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
func easy_excel_get_cell(handle int64, sheet, cell *C.zend_string, mode int64) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	v, err := wb.GetCell(goStr(sheet), goStr(cell), core.GetMode(mode))
	return pair(v, err)
}

//export_php:function easy_excel_read_rows(int $handle, string $sheet, int $startRow, int $maxRows, bool $raw): array
func easy_excel_read_rows(handle int64, sheet *C.zend_string, startRow, maxRows int64, raw bool) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return pair(nil, err)
	}
	rows, more, err := wb.ReadRows(goStr(sheet), int(startRow), int(maxRows), raw)
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
func easy_excel_set_number_format(handle int64, sheet, ref, code *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SetNumberFormat(goStr(sheet), goStr(ref), goStr(code)))
}

//export_php:function easy_excel_merge_cells(int $handle, string $sheet, string $range): ?string
func easy_excel_merge_cells(handle int64, sheet, ref *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.MergeCells(goStr(sheet), goStr(ref)))
}

// --- save -------------------------------------------------------------------------

//export_php:function easy_excel_save_xlsx(int $handle, string $path): ?string
func easy_excel_save_xlsx(handle int64, path *C.zend_string) unsafe.Pointer {
	wb, err := workbook(handle)
	if err != nil {
		return errOnly(err)
	}
	return errOnly(wb.SaveXlsx(goStr(path)))
}

//export_php:function easy_excel_save_csv(int $handle, string $path, string $sheet, string $delimiter, bool $crlf, bool $bom, bool $guardFormulas): ?string
func easy_excel_save_csv(handle int64, path, sheet, delimiter *C.zend_string, crlf, bom, guard bool) unsafe.Pointer {
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
