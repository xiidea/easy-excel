<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Exception;

/**
 * Immutable rectangular cell range, PhpSpreadsheet style (wave 5.2). A thin
 * value object over the existing Coordinate helpers — no extension
 * involvement. Corners are normalised so from() is always the top-left.
 */
final class AddressRange implements \Stringable
{
    private readonly CellAddress $from;

    private readonly CellAddress $to;

    public function __construct(CellAddress $from, CellAddress $to)
    {
        // Normalise so a range built "backwards" (D9:B2) still reads top-left
        // first, matching how Coordinate::rangeBoundaries treats it.
        $this->from = CellAddress::fromColumnAndRow(
            \min($from->columnId(), $to->columnId()),
            \min($from->rowId(), $to->rowId()),
        );
        $this->to = CellAddress::fromColumnAndRow(
            \max($from->columnId(), $to->columnId()),
            \max($from->rowId(), $to->rowId()),
        );
    }

    /** Build from an "A1:C10" style string. */
    public static function fromCellRange(string $range): self
    {
        $parts = \explode(':', $range);
        if (\count($parts) !== 2 || $parts[0] === '' || $parts[1] === '') {
            throw new Exception("easy-excel: invalid cell range \"$range\"");
        }

        return new self(
            CellAddress::fromCellAddress($parts[0]),
            CellAddress::fromCellAddress($parts[1]),
        );
    }

    public function from(): CellAddress
    {
        return $this->from;
    }

    public function to(): CellAddress
    {
        return $this->to;
    }

    /** The range without absolute markers, e.g. "B2:D9". */
    public function cellRange(): string
    {
        return $this->from->cellAddress() . ':' . $this->to->cellAddress();
    }

    /** The range with absolute markers, e.g. "$B$2:$D$9". */
    public function absoluteCellRange(): string
    {
        return $this->from->absoluteCellAddress() . ':' . $this->to->absoluteCellAddress();
    }

    public function __toString(): string
    {
        return $this->cellRange();
    }
}
