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

function easy_excel_open(string $path, string $password = ''): array
{
    EasyExcelFake::$log[] = ['open', [$path, $password]];

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

function easy_excel_read_rows(int $handle, string $sheet, int $startRow, int $maxRows, bool $raw, bool $calc = false): array
{
    EasyExcelFake::$log[] = ['read_rows', [$handle, $sheet, $startRow, $maxRows, $raw, $calc]];
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

function easy_excel_apply_style(int $handle, string $sheet, string $range, string $styleJson): ?string
{
    $spec = \json_decode($styleJson, true);
    if (!\is_array($spec)) {
        return 'fake: style spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['apply_style', [$handle, $sheet, $range, $spec]];

    return null;
}

function easy_excel_set_col_width(int $handle, string $sheet, int $startCol, int $endCol, float $width): ?string
{
    EasyExcelFake::$log[] = ['set_col_width', [$handle, $sheet, $startCol, $endCol, $width]];

    return null;
}

function easy_excel_set_col_autosize(int $handle, string $sheet, int $startCol, int $endCol): ?string
{
    EasyExcelFake::$log[] = ['set_col_autosize', [$handle, $sheet, $startCol, $endCol]];

    return null;
}

function easy_excel_set_row_height(int $handle, string $sheet, int $row, float $height): ?string
{
    EasyExcelFake::$log[] = ['set_row_height', [$handle, $sheet, $row, $height]];

    return null;
}

function easy_excel_freeze_panes(int $handle, string $sheet, string $topLeftCell): ?string
{
    EasyExcelFake::$log[] = ['freeze_panes', [$handle, $sheet, $topLeftCell]];

    return null;
}

function easy_excel_auto_filter(int $handle, string $sheet, string $range): ?string
{
    EasyExcelFake::$log[] = ['auto_filter', [$handle, $sheet, $range]];

    return null;
}

function easy_excel_set_hyperlink(int $handle, string $sheet, string $cell, string $url, string $tooltip): ?string
{
    EasyExcelFake::$log[] = ['set_hyperlink', [$handle, $sheet, $cell, $url, $tooltip]];

    return null;
}

function easy_excel_set_comment(int $handle, string $sheet, string $cell, string $author, string $text): ?string
{
    EasyExcelFake::$log[] = ['set_comment', [$handle, $sheet, $cell, $author, $text]];

    return null;
}

function easy_excel_defined_name(int $handle, string $name, string $refersTo, string $scopeSheet): ?string
{
    EasyExcelFake::$log[] = ['defined_name', [$handle, $name, $refersTo, $scopeSheet]];

    return null;
}

function easy_excel_page_setup(int $handle, string $sheet, string $orientation, int $paperSize, int $fitToWidth, int $fitToHeight): ?string
{
    EasyExcelFake::$log[] = ['page_setup', [$handle, $sheet, $orientation, $paperSize, $fitToWidth, $fitToHeight]];

    return null;
}

function easy_excel_set_validation(int $handle, string $sheet, string $range, string $validationJson): ?string
{
    $spec = \json_decode($validationJson, true);
    if (!\is_array($spec)) {
        return 'fake: validation spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['set_validation', [$handle, $sheet, $range, $spec]];

    return null;
}

function easy_excel_set_conditional(int $handle, string $sheet, string $range, string $rulesJson): ?string
{
    $rules = \json_decode($rulesJson, true);
    if (!\is_array($rules)) {
        return 'fake: conditional rules are not valid JSON';
    }
    EasyExcelFake::$log[] = ['set_conditional', [$handle, $sheet, $range, $rules]];

    return null;
}

function easy_excel_add_image(int $handle, string $sheet, string $cell, string $imageJson): ?string
{
    $spec = \json_decode($imageJson, true);
    if (!\is_array($spec)) {
        return 'fake: image spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['add_image', [$handle, $sheet, $cell, $spec]];

    return null;
}

function easy_excel_protect_sheet(int $handle, string $sheet, string $protectionJson): ?string
{
    $spec = \json_decode($protectionJson, true);
    if (!\is_array($spec)) {
        return 'fake: protection spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['protect_sheet', [$handle, $sheet, $spec]];

    return null;
}

function easy_excel_add_chart(int $handle, string $sheet, string $cell, string $chartJson): ?string
{
    $spec = \json_decode($chartJson, true);
    if (!\is_array($spec)) {
        return 'fake: chart spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['add_chart', [$handle, $sheet, $cell, $spec]];

    return null;
}

function easy_excel_save_xlsx(int $handle, string $path, string $password = ''): ?string
{
    EasyExcelFake::$log[] = ['save_xlsx', [$handle, $path, $password]];
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

function easy_excel_doc_props(int $handle, string $propsJson): ?string
{
    $props = \json_decode($propsJson, true);
    if (!\is_array($props)) {
        return 'fake: doc props are not valid JSON';
    }
    EasyExcelFake::$log[] = ['doc_props', [$handle, $props]];

    return null;
}

function easy_excel_custom_prop(int $handle, string $propJson): ?string
{
    $prop = \json_decode($propJson, true);
    if (!\is_array($prop)) {
        return 'fake: custom property is not valid JSON';
    }
    EasyExcelFake::$log[] = ['custom_prop', [$handle, $prop]];

    return null;
}

function easy_excel_unmerge_cells(int $handle, string $sheet, string $range): ?string
{
    EasyExcelFake::$log[] = ['unmerge_cells', [$handle, $sheet, $range]];

    return null;
}

function easy_excel_get_merges(int $handle, string $sheet): array
{
    $merges = [];
    foreach (EasyExcelFake::$log as [$fn, $args]) {
        if ($fn === 'merge_cells' && $args[0] === $handle && $args[1] === $sheet) {
            $merges[$args[2]] = true;
        }
        if ($fn === 'unmerge_cells' && $args[0] === $handle && $args[1] === $sheet) {
            unset($merges[$args[2]]);
        }
    }

    return [\array_keys($merges), null];
}

function easy_excel_get_style(int $handle, string $sheet, string $cell): array
{
    // fold every apply_style / set_number_format call whose range contains
    // the cell, in order — close enough to the Go fold for shim tests
    [$col, $row] = \EasyExcel\Compat\Cell\Coordinate::indexesFromString($cell);
    $contains = static function (string $range) use ($col, $row): bool {
        if (\preg_match('/^[A-Z]+(:[A-Z]+)?$/', $range)) { // full column(s)
            $parts = \explode(':', $range);
            $c1 = \EasyExcel\Compat\Cell\Coordinate::columnIndexFromString($parts[0]);
            $c2 = \EasyExcel\Compat\Cell\Coordinate::columnIndexFromString($parts[1] ?? $parts[0]);

            return $col >= \min($c1, $c2) && $col <= \max($c1, $c2);
        }
        [[$c1, $r1], [$c2, $r2]] = \EasyExcel\Compat\Cell\Coordinate::rangeBoundaries($range);

        return $col >= $c1 && $col <= $c2 && $row >= $r1 && $row <= $r2;
    };
    $merge = static function (array $base, array $patch) use (&$merge): array {
        foreach ($patch as $k => $v) {
            $base[$k] = \is_array($v) && \is_array($base[$k] ?? null) ? $merge($base[$k], $v) : $v;
        }

        return $base;
    };
    $spec = [];
    foreach (EasyExcelFake::$log as [$fn, $args]) {
        if ($fn === 'set_default_style' && $args[0] === $handle) {
            $spec = $merge($args[1], $spec);
        }
        if ($fn === 'apply_style' && $args[0] === $handle && $args[1] === $sheet && $contains($args[2])) {
            $spec = $merge($spec, $args[3]);
        }
        if ($fn === 'set_number_format' && $args[0] === $handle && $args[1] === $sheet && $contains($args[2])) {
            $spec = $merge($spec, ['numberFormat' => ['formatCode' => $args[3]]]);
        }
    }

    return [\json_encode($spec), null];
}

function easy_excel_get_validations(int $handle, string $sheet): array
{
    $out = [];
    foreach (EasyExcelFake::calls('set_validation') as [$fn, $args]) {
        if ($args[0] === $handle && $args[1] === $sheet) {
            $out[] = ['sqref' => $args[2], 'spec' => $args[3]];
        }
    }

    return [\json_encode($out), null];
}

function easy_excel_get_conditionals(int $handle, string $sheet): array
{
    $out = [];
    foreach (EasyExcelFake::calls('set_conditional') as [$fn, $args]) {
        if ($args[0] === $handle && $args[1] === $sheet) {
            $out[$args[2]] = $args[3];
        }
    }

    return [\json_encode($out), null];
}

function easy_excel_get_defined_names(int $handle): array
{
    $out = [];
    foreach (EasyExcelFake::calls('defined_name') as [$fn, $args]) {
        if ($args[0] === $handle) {
            $out[] = ['name' => $args[1], 'refersTo' => $args[2], 'scope' => $args[3]];
        }
    }

    return [\json_encode($out), null];
}

function easy_excel_set_default_style(int $handle, string $styleJson): ?string
{
    $spec = \json_decode($styleJson, true);
    if (!\is_array($spec)) {
        return 'fake: default style is not valid JSON';
    }
    EasyExcelFake::$log[] = ['set_default_style', [$handle, $spec]];

    return null;
}

function easy_excel_insert_rows(int $handle, string $sheet, int $row, int $count): ?string
{
    EasyExcelFake::$log[] = ['insert_rows', [$handle, $sheet, $row, $count]];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    $cells = [];
    foreach ($s['cells'] as $r => $cols) {
        $cells[$r >= $row ? $r + $count : $r] = $cols;
    }
    $s['cells'] = $cells;
    $s['maxRow'] += $count;

    return null;
}

function easy_excel_remove_rows(int $handle, string $sheet, int $row, int $count): ?string
{
    EasyExcelFake::$log[] = ['remove_rows', [$handle, $sheet, $row, $count]];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    $cells = [];
    foreach ($s['cells'] as $r => $cols) {
        if ($r >= $row && $r < $row + $count) {
            continue;
        }
        $cells[$r >= $row + $count ? $r - $count : $r] = $cols;
    }
    $s['cells'] = $cells;
    $s['maxRow'] = \max(0, $s['maxRow'] - $count);

    return null;
}

function easy_excel_insert_cols(int $handle, string $sheet, int $col, int $count): ?string
{
    EasyExcelFake::$log[] = ['insert_cols', [$handle, $sheet, $col, $count]];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    foreach ($s['cells'] as $r => $cols) {
        $shifted = [];
        foreach ($cols as $c => $v) {
            $shifted[$c >= $col ? $c + $count : $c] = $v;
        }
        $s['cells'][$r] = $shifted;
    }
    $s['maxCol'] += $count;

    return null;
}

function easy_excel_remove_cols(int $handle, string $sheet, int $col, int $count): ?string
{
    EasyExcelFake::$log[] = ['remove_cols', [$handle, $sheet, $col, $count]];
    $s = &EasyExcelFake::sheet($handle, $sheet);
    foreach ($s['cells'] as $r => $cols) {
        $shifted = [];
        foreach ($cols as $c => $v) {
            if ($c >= $col && $c < $col + $count) {
                continue;
            }
            $shifted[$c >= $col + $count ? $c - $count : $c] = $v;
        }
        $s['cells'][$r] = $shifted;
    }
    $s['maxCol'] = \max(0, $s['maxCol'] - $count);

    return null;
}

function easy_excel_move_sheet(int $handle, string $sheet, int $index): ?string
{
    EasyExcelFake::$log[] = ['move_sheet', [$handle, $sheet, $index]];
    $wb = &EasyExcelFake::$store[$handle];
    $order = \array_values(\array_diff($wb['order'], [$sheet]));
    \array_splice($order, $index, 0, [$sheet]);
    $wb['order'] = $order;

    return null;
}

function easy_excel_copy_sheet(int $handle, string $from, string $newName): array
{
    EasyExcelFake::$log[] = ['copy_sheet', [$handle, $from, $newName]];
    $wb = &EasyExcelFake::$store[$handle];
    $wb['sheets'][$newName] = $wb['sheets'][$from];
    $wb['order'][] = $newName;

    return [\count($wb['order']) - 1, null];
}

function easy_excel_sheet_view(int $handle, string $sheet, string $viewJson): ?string
{
    $spec = \json_decode($viewJson, true);
    if (!\is_array($spec)) {
        return 'fake: sheet view spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['sheet_view', [$handle, $sheet, $spec]];

    return null;
}

function easy_excel_header_footer(int $handle, string $sheet, string $hfJson): ?string
{
    $spec = \json_decode($hfJson, true);
    if (!\is_array($spec)) {
        return 'fake: header/footer spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['header_footer', [$handle, $sheet, $spec]];

    return null;
}

function easy_excel_page_margins(int $handle, string $sheet, string $marginsJson): ?string
{
    $spec = \json_decode($marginsJson, true);
    if (!\is_array($spec)) {
        return 'fake: margins spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['page_margins', [$handle, $sheet, $spec]];

    return null;
}

function easy_excel_set_rich_text(int $handle, string $sheet, string $cell, string $runsJson): ?string
{
    $runs = \json_decode($runsJson, true);
    if (!\is_array($runs)) {
        return 'fake: rich text is not valid JSON';
    }
    EasyExcelFake::$log[] = ['set_rich_text', [$handle, $sheet, $cell, $runs]];

    return null;
}

function easy_excel_add_image_bytes(int $handle, string $sheet, string $cell, string $imageJson): ?string
{
    $spec = \json_decode($imageJson, true);
    if (!\is_array($spec)) {
        return 'fake: image spec is not valid JSON';
    }
    EasyExcelFake::$log[] = ['add_image_bytes', [$handle, $sheet, $cell, $spec]];

    return null;
}

function easy_excel_auto_filter_columns(int $handle, string $sheet, string $range, string $columnsJson): ?string
{
    $cols = \json_decode($columnsJson, true);
    if (!\is_array($cols)) {
        return 'fake: auto-filter columns are not valid JSON';
    }
    EasyExcelFake::$log[] = ['auto_filter_columns', [$handle, $sheet, $range, $cols]];

    return null;
}

function easy_excel_stats(): array
{
    return [\count(EasyExcelFake::$store), 0];
}
