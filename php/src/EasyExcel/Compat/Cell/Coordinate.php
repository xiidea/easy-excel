<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Exception;

/**
 * Pure-PHP port of PhpSpreadsheet's Coordinate helpers (no extension calls:
 * coordinate math must never cross the CGO boundary, PLAN.md C2).
 */
abstract class Coordinate
{
    public const A1_COORDINATE_REGEX = '/^(?<absolute_col>\$?)(?<col_ref>[A-Z]{1,3})(?<absolute_row>\$?)(?<row_ref>\d{1,7})$/i';

    /** @var array<string, int> */
    private static array $columnLookup = [];

    /** @return array{0: string, 1: string} [column, row], with $ kept */
    public static function coordinateFromString(string $coordinates): array
    {
        if (\preg_match(self::A1_COORDINATE_REGEX, $coordinates, $m)) {
            return [$m['absolute_col'] . \strtoupper($m['col_ref']), $m['absolute_row'] . $m['row_ref']];
        }
        if (\str_contains($coordinates, ':')) {
            throw new Exception('Cell coordinate string can not be a range of cells');
        }
        if ($coordinates === '') {
            throw new Exception('Cell coordinate can not be zero-length string');
        }

        throw new Exception('Invalid cell coordinate ' . $coordinates);
    }

    /** @return array{0: int, 1: int} [columnIndex (1-based), row] */
    public static function indexesFromString(string $coordinates): array
    {
        [$col, $row] = self::coordinateFromString($coordinates);

        return [self::columnIndexFromString(\ltrim($col, '$')), (int) \ltrim($row, '$')];
    }

    public static function columnIndexFromString(string $columnAddress): int
    {
        $columnAddress = \strtoupper($columnAddress);
        if (isset(self::$columnLookup[$columnAddress])) {
            return self::$columnLookup[$columnAddress];
        }
        $len = \strlen($columnAddress);
        if ($len === 0 || $len > 3 || !\preg_match('/^[A-Z]+$/', $columnAddress)) {
            throw new Exception('Column string index can not be ' . ($len ? 'longer than 3 characters' : 'empty'));
        }
        $index = 0;
        for ($i = 0; $i < $len; ++$i) {
            $index = $index * 26 + (\ord($columnAddress[$i]) - 64);
        }
        self::$columnLookup[$columnAddress] = $index;

        return $index;
    }

    public static function stringFromColumnIndex(int $columnIndex): string
    {
        if ($columnIndex < 1) {
            throw new Exception('Column index has to be a positive integer');
        }
        $column = '';
        while ($columnIndex > 0) {
            $columnIndex--;
            $column = \chr(65 + ($columnIndex % 26)) . $column;
            $columnIndex = \intdiv($columnIndex, 26);
        }

        return $column;
    }

    /** @return array{0: array{0: int, 1: int}, 1: array{0: int, 1: int}} [[startCol, startRow], [endCol, endRow]] */
    public static function rangeBoundaries(string $range): array
    {
        $range = \str_replace('$', '', \strtoupper($range));
        if (!\str_contains($range, ':')) {
            $range .= ':' . $range;
        }
        [$start, $end] = \explode(':', $range, 2);
        [$startCol, $startRow] = self::indexesFromString($start);
        [$endCol, $endRow] = self::indexesFromString($end);

        return [[$startCol, $startRow], [$endCol, $endRow]];
    }

    /** @return array{0: int, 1: int} [columns, rows] */
    public static function rangeDimension(string $range): array
    {
        [[$sc, $sr], [$ec, $er]] = self::rangeBoundaries($range);

        return [$ec - $sc + 1, $er - $sr + 1];
    }

    /** @return list<array{0: string, 1?: string}> */
    public static function splitRange(string $range): array
    {
        $exploded = \explode(',', $range);
        $out = [];
        foreach ($exploded as $value) {
            $out[] = \explode(':', $value);
        }

        return $out;
    }
}
