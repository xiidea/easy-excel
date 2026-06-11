<?php

declare(strict_types=1);

/*
 * In-memory fake of the easy_excel_* extension ABI, mirroring the bridge's
 * return conventions and the Go compat binder closely enough to unit-test
 * the PHP shim without FrankenPHP. Every call is recorded in
 * EasyExcelFake::$log so tests can assert batching behavior.
 */

final class EasyExcelFake
{
    /** @var array<int, array{sheets: array<string, array{cells: array<int, array<int, mixed>>, maxRow: int, maxCol: int}>, order: list<string>, active: int}> */
    public static array $store = [];

    /** @var list<array{0: string, 1: array}> */
    public static array $log = [];

    public static int $nextHandle = 1;

    public static function reset(): void
    {
        self::$store = [];
        self::$log = [];
        self::$nextHandle = 1;
    }

    /** @return list<array> */
    public static function calls(string $fn): array
    {
        return \array_values(\array_filter(self::$log, static fn (array $e): bool => $e[0] === $fn));
    }

    /** Mirrors extension/compat/values.go Decode(). */
    public static function bind(mixed $v): mixed
    {
        if (\is_array($v)) { // explicit [marker, value]
            [$marker, $value] = $v;

            return match ($marker) {
                '=s' => (string) $value,
                '=n' => (float) $value,
                '=b' => (bool) $value,
                '=f' => '=' . \ltrim((string) $value, '='),
                default => throw new \LogicException("fake: unknown marker $marker"),
            };
        }
        if (\is_string($v)) {
            if (\strlen($v) > 1 && $v[0] === '=') {
                return $v; // formula, kept with '='
            }
            if (\preg_match('/^[+-]?(\d+\.?\d*|\d*\.?\d+)([eE][+-]?\d+)?$/', $v)) {
                $digits = \ltrim($v, '+-');
                if (!(\strlen($digits) > 1 && $digits[0] === '0' && $digits[1] !== '.')) {
                    return (float) $v;
                }
            }

            return $v;
        }
        if (\is_int($v)) {
            return (float) $v;
        }

        return $v; // float, bool, null
    }

    public static function &sheet(int $handle, string $name): array
    {
        if (!isset(self::$store[$handle])) {
            throw new \LogicException('fake: bad handle');
        }
        if (!isset(self::$store[$handle]['sheets'][$name])) {
            throw new \LogicException("fake: no sheet $name");
        }

        return self::$store[$handle]['sheets'][$name];
    }

    public static function stringify(mixed $v): string
    {
        if (\is_bool($v)) {
            return $v ? '1' : '0';
        }
        if (\is_float($v)) {
            return (string) ($v == (int) $v ? (int) $v : $v);
        }

        return (string) $v;
    }
}

function easy_excel_version(): string
{
    return '0.1.0-fake';
}

function easy_excel_new(): array
{
    $h = EasyExcelFake::$nextHandle++;
    EasyExcelFake::$store[$h] = [
        'sheets' => ['Worksheet' => ['cells' => [], 'maxRow' => 0, 'maxCol' => 0]],
        'order' => ['Worksheet'],
        'active' => 0,
    ];
    EasyExcelFake::$log[] = ['new', [$h]];

    return [$h, null];
}

function easy_excel_open(string $path): array
{
    return [null, "fake: open not supported"];
}

function easy_excel_close(int $handle): ?string
{
    unset(EasyExcelFake::$store[$handle]);
    EasyExcelFake::$log[] = ['close', [$handle]];

    return null;
}

function easy_excel_add_sheet(int $handle, string $name): array
{
    if (isset(EasyExcelFake::$store[$handle]['sheets'][$name])) {
        return [null, "sheet $name already exists"];
    }
    EasyExcelFake::$store[$handle]['sheets'][$name] = ['cells' => [], 'maxRow' => 0, 'maxCol' => 0];
    EasyExcelFake::$store[$handle]['order'][] = $name;

    return [\count(EasyExcelFake::$store[$handle]['order']) - 1, null];
}

function easy_excel_delete_sheet(int $handle, string $name): ?string
{
    $wb = &EasyExcelFake::$store[$handle];
    if (\count($wb['order']) === 1) {
        return 'cannot remove the only sheet';
    }
    unset($wb['sheets'][$name]);
    $wb['order'] = \array_values(\array_diff($wb['order'], [$name]));

    return null;
}

function easy_excel_rename_sheet(int $handle, string $old, string $new): ?string
{
    $wb = &EasyExcelFake::$store[$handle];
    if (!isset($wb['sheets'][$old])) {
        return "no sheet $old";
    }
    $wb['sheets'][$new] = $wb['sheets'][$old];
    unset($wb['sheets'][$old]);
    $wb['order'] = \array_map(static fn (string $n): string => $n === $old ? $new : $n, $wb['order']);

    return null;
}

function easy_excel_sheets(int $handle): array
{
    return [EasyExcelFake::$store[$handle]['order'], null];
}

function easy_excel_set_active_sheet(int $handle, int $index): ?string
{
    EasyExcelFake::$store[$handle]['active'] = $index;

    return null;
}

function easy_excel_active_sheet(int $handle): array
{
    $wb = EasyExcelFake::$store[$handle];

    return [[$wb['active'], $wb['order'][$wb['active']]], null];
}

function easy_excel_write_rows(int $handle, string $sheet, int $startRow, int $startCol, array $rows): ?string
{
    EasyExcelFake::$log[] = ['write_rows', [$handle, $sheet, $startRow, $startCol, \count($rows)]];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    foreach ($rows as $i => $cols) {
        foreach ($cols as $j => $v) {
            if ($v === null) {
                continue;
            }
            $row = $startRow + $i;
            $col = $startCol + $j;
            $s['cells'][$row][$col] = EasyExcelFake::bind($v);
            $s['maxRow'] = \max($s['maxRow'], $row);
            $s['maxCol'] = \max($s['maxCol'], $col);
        }
    }

    return null;
}

function easy_excel_set_cell(int $handle, string $sheet, string $cell, array $value): ?string
{
    [$col, $row] = \EasyExcel\Compat\Cell\Coordinate::indexesFromString($cell);
    $encoded = \count($value) === 2 ? $value : $value[0];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    $s['cells'][$row][$col] = EasyExcelFake::bind($encoded);
    $s['maxRow'] = \max($s['maxRow'], $row);
    $s['maxCol'] = \max($s['maxCol'], $col);

    return null;
}

function easy_excel_get_cell(int $handle, string $sheet, string $cell, int $mode): array
{
    [$col, $row] = \EasyExcel\Compat\Cell\Coordinate::indexesFromString($cell);
    $s = &EasyExcelFake::sheet($handle, $sheet);
    $v = $s['cells'][$row][$col] ?? null;
    if ($mode === 1 && $v !== null) { // formatted
        $v = EasyExcelFake::stringify($v);
    }
    if ($mode === 2) { // calculated: not supported by the fake
        $v = \is_string($v) && \str_starts_with($v, '=') ? '#FAKE!' : $v;
    }

    return [$v, null];
}

function easy_excel_read_rows(int $handle, string $sheet, int $startRow, int $maxRows, bool $raw): array
{
    EasyExcelFake::$log[] = ['read_rows', [$handle, $sheet, $startRow, $maxRows, $raw]];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    $out = [];
    $last = \min($s['maxRow'], $startRow + $maxRows - 1);
    for ($r = $startRow; $r <= $last; ++$r) {
        $cols = [];
        $rowCells = $s['cells'][$r] ?? [];
        $width = $rowCells === [] ? 0 : \max(\array_keys($rowCells));
        for ($c = 1; $c <= $width; ++$c) {
            $cols[] = isset($rowCells[$c]) ? EasyExcelFake::stringify($rowCells[$c]) : '';
        }
        $out[] = $cols;
    }

    return [[$out, $last < $s['maxRow']], null];
}

function easy_excel_dimensions(int $handle, string $sheet): array
{
    $s = &EasyExcelFake::sheet($handle, $sheet);

    return [[$s['maxRow'], $s['maxCol']], null];
}

function easy_excel_set_number_format(int $handle, string $sheet, string $range, string $code): ?string
{
    EasyExcelFake::$log[] = ['set_number_format', [$handle, $sheet, $range, $code]];

    return null;
}

function easy_excel_merge_cells(int $handle, string $sheet, string $range): ?string
{
    EasyExcelFake::$log[] = ['merge_cells', [$handle, $sheet, $range]];

    return null;
}

function easy_excel_save_xlsx(int $handle, string $path): ?string
{
    EasyExcelFake::$log[] = ['save_xlsx', [$handle, $path]];
    \file_put_contents($path, \json_encode(EasyExcelFake::$store[$handle]));

    return null;
}

function easy_excel_save_csv(int $handle, string $path, string $sheet, string $delimiter, bool $crlf, bool $bom, bool $guard): ?string
{
    $s = &EasyExcelFake::sheet($handle, $sheet);
    $fh = \fopen($path, 'wb');
    if ($bom) {
        \fwrite($fh, "\xEF\xBB\xBF");
    }
    for ($r = 1; $r <= $s['maxRow']; ++$r) {
        $cols = [];
        $rowCells = $s['cells'][$r] ?? [];
        $width = $rowCells === [] ? 0 : \max(\array_keys($rowCells));
        for ($c = 1; $c <= $width; ++$c) {
            $v = isset($rowCells[$c]) ? EasyExcelFake::stringify($rowCells[$c]) : '';
            if ($guard && $v !== '' && \str_contains('=+-@', $v[0])) {
                $v = "'" . $v;
            }
            $cols[] = $v;
        }
        \fputcsv($fh, $cols, $delimiter, '"', '', $crlf ? "\r\n" : "\n");
    }
    \fclose($fh);

    return null;
}

function easy_excel_stats(): array
{
    return [\count(EasyExcelFake::$store), 0];
}
