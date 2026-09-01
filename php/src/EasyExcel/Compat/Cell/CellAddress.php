<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Exception;

/**
 * Immutable single-cell address, PhpSpreadsheet style (wave 5.2). A thin value
 * object over the existing Coordinate helpers — no extension involvement.
 */
final class CellAddress implements \Stringable
{
    public function __construct(
        private readonly string $columnName,
        private readonly int $columnId,
        private readonly int $rowId,
    ) {
    }

    /** Build from an "A1" style string. */
    public static function fromCellAddress(string $cellAddress): self
    {
        [$column, $row] = Coordinate::coordinateFromString($cellAddress);

        return new self($column, Coordinate::columnIndexFromString($column), (int) $row);
    }

    /** Build from 1-based column and row indexes. */
    public static function fromColumnAndRow(int $columnId, int $rowId): self
    {
        if ($columnId < 1 || $rowId < 1) {
            throw new Exception('easy-excel: column and row indexes are 1-based');
        }

        return new self(Coordinate::stringFromColumnIndex($columnId), $columnId, $rowId);
    }

    public function columnName(): string
    {
        return $this->columnName;
    }

    public function columnId(): int
    {
        return $this->columnId;
    }

    public function rowId(): int
    {
        return $this->rowId;
    }

    /** The address without absolute markers, e.g. "B7". */
    public function cellAddress(): string
    {
        return $this->columnName . $this->rowId;
    }

    /** The address with absolute markers, e.g. "$B$7". */
    public function absoluteCellAddress(): string
    {
        return '$' . $this->columnName . '$' . $this->rowId;
    }

    public function nextRow(int $offset = 1): self
    {
        return self::fromColumnAndRow($this->columnId, \max(1, $this->rowId + $offset));
    }

    public function previousRow(int $offset = 1): self
    {
        return $this->nextRow(-$offset);
    }

    public function nextColumn(int $offset = 1): self
    {
        return self::fromColumnAndRow(\max(1, $this->columnId + $offset), $this->rowId);
    }

    public function previousColumn(int $offset = 1): self
    {
        return $this->nextColumn(-$offset);
    }

    public function __toString(): string
    {
        return $this->cellAddress();
    }
}
