<?php

declare(strict_types=1);

namespace EasyExcel;

use EasyExcel\Exception\BadHandle;
use EasyExcel\Exception\EasyExcelException;
use EasyExcel\Exception\Overloaded;
use EasyExcel\Exception\PathDenied;

/**
 * Thin typed wrapper over the flat easy_excel_* extension ABI.
 *
 * Decodes the bridge's two return conventions — `?string` (null = success)
 * and `array{0: mixed, 1: ?string}` — and converts error-message prefixes
 * into typed exceptions. Every other class in this package talks to the
 * extension exclusively through here.
 */
final class Native
{
    public const GET_RAW = 0;
    public const GET_FORMATTED = 1;
    public const GET_CALCULATED = 2;

    /** Explicit cell markers; must match extension/compat/values.go. */
    public const MARK_STRING = '=s';
    public const MARK_NUMERIC = '=n';
    public const MARK_BOOL = '=b';
    public const MARK_FORMULA = '=f';

    private function __construct()
    {
    }

    public static function available(): bool
    {
        return \function_exists('easy_excel_new');
    }

    public static function assertAvailable(): void
    {
        if (!self::available()) {
            throw new EasyExcelException(
                'The easy_excel extension is not loaded. Build FrankenPHP with '
                . '--with github.com/ronisaha/easy-excel/extension/build (see README).'
            );
        }
    }

    public static function newWorkbook(): int
    {
        self::assertAvailable();

        return (int) self::unwrap(\easy_excel_new());
    }

    public static function open(string $path): int
    {
        self::assertAvailable();

        return (int) self::unwrap(\easy_excel_open($path));
    }

    public static function close(int $handle): void
    {
        if (self::available()) {
            self::check(\easy_excel_close($handle));
        }
    }

    public static function addSheet(int $handle, string $name): int
    {
        return (int) self::unwrap(\easy_excel_add_sheet($handle, $name));
    }

    public static function deleteSheet(int $handle, string $name): void
    {
        self::check(\easy_excel_delete_sheet($handle, $name));
    }

    public static function renameSheet(int $handle, string $old, string $new): void
    {
        self::check(\easy_excel_rename_sheet($handle, $old, $new));
    }

    /** @return list<string> */
    public static function sheets(int $handle): array
    {
        return self::unwrap(\easy_excel_sheets($handle));
    }

    public static function setActiveSheet(int $handle, int $index): void
    {
        self::check(\easy_excel_set_active_sheet($handle, $index));
    }

    /** @return array{0: int, 1: string} [position, name] */
    public static function activeSheet(int $handle): array
    {
        return self::unwrap(\easy_excel_active_sheet($handle));
    }

    /**
     * Hot path: writes a batch of rows anchored at (startRow, startCol).
     *
     * @param list<list<mixed>> $rows scalars, or [marker, value] pairs
     */
    public static function writeRows(int $handle, string $sheet, int $startRow, int $startCol, array $rows): void
    {
        self::check(\easy_excel_write_rows($handle, $sheet, $startRow, $startCol, $rows));
    }

    public static function setCell(int $handle, string $sheet, string $cell, mixed $value, ?string $marker = null): void
    {
        $encoded = $marker === null ? [$value] : [$marker, $value];
        self::check(\easy_excel_set_cell($handle, $sheet, $cell, $encoded));
    }

    public static function getCell(int $handle, string $sheet, string $cell, int $mode): mixed
    {
        return self::unwrap(\easy_excel_get_cell($handle, $sheet, $cell, $mode));
    }

    /** @return array{0: list<list<string>>, 1: bool} [rows, more] */
    public static function readRows(int $handle, string $sheet, int $startRow, int $maxRows, bool $raw, bool $calc = false): array
    {
        return self::unwrap(\easy_excel_read_rows($handle, $sheet, $startRow, $maxRows, $raw, $calc));
    }

    /** @return array{0: int, 1: int} [highestRow, highestCol] */
    public static function dimensions(int $handle, string $sheet): array
    {
        return self::unwrap(\easy_excel_dimensions($handle, $sheet));
    }

    public static function setNumberFormat(int $handle, string $sheet, string $range, string $code): void
    {
        self::check(\easy_excel_set_number_format($handle, $sheet, $range, $code));
    }

    public static function mergeCells(int $handle, string $sheet, string $range): void
    {
        self::check(\easy_excel_merge_cells($handle, $sheet, $range));
    }

    /** @param array<string, array<string, mixed>> $spec PhpSpreadsheet applyFromArray shape */
    public static function applyStyle(int $handle, string $sheet, string $range, array $spec): void
    {
        self::check(\easy_excel_apply_style($handle, $sheet, $range, \json_encode($spec, \JSON_THROW_ON_ERROR)));
    }

    public static function setColWidth(int $handle, string $sheet, int $startCol, int $endCol, float $width): void
    {
        self::check(\easy_excel_set_col_width($handle, $sheet, $startCol, $endCol, $width));
    }

    public static function setColAutoSize(int $handle, string $sheet, int $startCol, int $endCol): void
    {
        self::check(\easy_excel_set_col_autosize($handle, $sheet, $startCol, $endCol));
    }

    public static function setRowHeight(int $handle, string $sheet, int $row, float $height): void
    {
        self::check(\easy_excel_set_row_height($handle, $sheet, $row, $height));
    }

    public static function freezePanes(int $handle, string $sheet, string $topLeftCell): void
    {
        self::check(\easy_excel_freeze_panes($handle, $sheet, $topLeftCell));
    }

    public static function autoFilter(int $handle, string $sheet, string $range): void
    {
        self::check(\easy_excel_auto_filter($handle, $sheet, $range));
    }

    public static function setHyperlink(int $handle, string $sheet, string $cell, string $url, string $tooltip): void
    {
        self::check(\easy_excel_set_hyperlink($handle, $sheet, $cell, $url, $tooltip));
    }

    public static function setComment(int $handle, string $sheet, string $cell, string $author, string $text): void
    {
        self::check(\easy_excel_set_comment($handle, $sheet, $cell, $author, $text));
    }

    public static function definedName(int $handle, string $name, string $refersTo, string $scopeSheet): void
    {
        self::check(\easy_excel_defined_name($handle, $name, $refersTo, $scopeSheet));
    }

    public static function pageSetup(int $handle, string $sheet, string $orientation, int $paperSize, int $fitToWidth, int $fitToHeight): void
    {
        self::check(\easy_excel_page_setup($handle, $sheet, $orientation, $paperSize, $fitToWidth, $fitToHeight));
    }

    /** @param array<string, mixed> $spec PhpSpreadsheet DataValidation shape */
    public static function setValidation(int $handle, string $sheet, string $range, array $spec): void
    {
        self::check(\easy_excel_set_validation($handle, $sheet, $range, \json_encode($spec, \JSON_THROW_ON_ERROR)));
    }

    /** @param list<array<string, mixed>> $rules conditional-formatting rules */
    public static function setConditional(int $handle, string $sheet, string $range, array $rules): void
    {
        self::check(\easy_excel_set_conditional($handle, $sheet, $range, \json_encode($rules, \JSON_THROW_ON_ERROR)));
    }

    /** @param array<string, mixed> $spec path/name/offsets/width/height */
    public static function addImage(int $handle, string $sheet, string $cell, array $spec): void
    {
        self::check(\easy_excel_add_image($handle, $sheet, $cell, \json_encode($spec, \JSON_THROW_ON_ERROR)));
    }

    /** @param array<string, mixed> $spec PhpSpreadsheet Worksheet\Protection flags */
    public static function protectSheet(int $handle, string $sheet, array $spec): void
    {
        self::check(\easy_excel_protect_sheet($handle, $sheet, \json_encode($spec, \JSON_THROW_ON_ERROR)));
    }

    /**
     * easy-excel native chart API (PhpSpreadsheet's chart object model is not
     * mapped; see COMPAT.md).
     *
     * @param array<string, mixed> $spec ['type', 'series' => [['name','categories','values']], 'title', 'legend' => ['position'], 'width', 'height']
     */
    public static function addChart(int $handle, string $sheet, string $cell, array $spec): void
    {
        self::check(\easy_excel_add_chart($handle, $sheet, $cell, \json_encode($spec, \JSON_THROW_ON_ERROR)));
    }

    public static function saveXlsx(int $handle, string $path): void
    {
        self::check(\easy_excel_save_xlsx($handle, $path));
    }

    public static function saveCsv(
        int $handle,
        string $path,
        string $sheet,
        string $delimiter = ',',
        bool $crlf = false,
        bool $bom = false,
        bool $guardFormulas = false,
    ): void {
        self::check(\easy_excel_save_csv($handle, $path, $sheet, $delimiter, $crlf, $bom, $guardFormulas));
    }

    // Success is null OR '': the generated bridge marshals Go's nil
    // unsafe.Pointer for a ?string return as an empty PHP string, and error
    // messages are never empty.

    private static function unwrap(array $result): mixed
    {
        $error = $result[1] ?? null;
        if ($error !== null && $error !== '') {
            self::raise((string) $error);
        }

        return $result[0];
    }

    private static function check(?string $error): void
    {
        if ($error !== null && $error !== '') {
            self::raise($error);
        }
    }

    private static function raise(string $message): never
    {
        if (\str_starts_with($message, 'OVERLOADED:')) {
            throw new Overloaded($message);
        }
        if (\str_starts_with($message, 'DENIED:')) {
            throw new PathDenied($message);
        }
        if (\str_starts_with($message, 'BADHANDLE:')) {
            throw new BadHandle($message);
        }

        throw new EasyExcelException($message);
    }
}
