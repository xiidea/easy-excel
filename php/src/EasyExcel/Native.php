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
    public static function readRows(int $handle, string $sheet, int $startRow, int $maxRows, bool $raw): array
    {
        return self::unwrap(\easy_excel_read_rows($handle, $sheet, $startRow, $maxRows, $raw));
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

    private static function unwrap(array $result): mixed
    {
        if (($result[1] ?? null) !== null) {
            self::raise((string) $result[1]);
        }

        return $result[0];
    }

    private static function check(?string $error): void
    {
        if ($error !== null) {
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
