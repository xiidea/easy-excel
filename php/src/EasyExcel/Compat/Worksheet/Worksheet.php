<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Cell\Cell;
use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Compat\Cell\DataType;
use EasyExcel\Compat\Cell\Hyperlink;
use EasyExcel\Compat\Comment;
use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Shared\Date;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Style\Style;
use EasyExcel\Native;

/**
 * PhpSpreadsheet-compatible worksheet over the extension's flat ABI.
 *
 * Per-cell writes land in a write-behind buffer and cross the CGO boundary
 * in row batches (PLAN.md C2): one extension call per FLUSH_AFTER_ROWS rows
 * instead of one per cell. The buffer is flushed before any read, save, or
 * structural change. Sequential writes therefore reach the Go StreamWriter
 * in ascending order and keep the workbook in constant-memory streaming mode.
 */
class Worksheet
{
    /** Buffered rows before an automatic flush. */
    public const FLUSH_AFTER_ROWS = 512;

    /** @var array<int, array<int, mixed>> row => col => scalar|[marker, value] */
    private array $buffer = [];

    private int $bufferedRows = 0;

    private ?PageSetup $pageSetup = null;

    public function __construct(private Spreadsheet $parent, private string $title)
    {
    }

    public function getParent(): Spreadsheet
    {
        return $this->parent;
    }

    public function getTitle(): string
    {
        return $this->title;
    }

    public function setTitle(string $title, bool $updateFormulaCellReferences = true, bool $validate = true): static
    {
        if ($title === $this->title) {
            return $this;
        }
        Native::renameSheet($this->parent->getHandle(), $this->title, $title);
        $this->title = $title;

        return $this;
    }

    // --- writes ---------------------------------------------------------------

    public function setCellValue(string|array $coordinate, mixed $value): static
    {
        [$col, $row] = $this->toIndexes($coordinate);
        $this->bufferCell($row, $col, $this->bindValue($value));

        return $this;
    }

    public function setCellValueByColumnAndRow(int $columnIndex, int $row, mixed $value): static
    {
        $this->bufferCell($row, $columnIndex, $this->bindValue($value));

        return $this;
    }

    public function setCellValueExplicit(string|array $coordinate, mixed $value, string $dataType = DataType::TYPE_STRING): static
    {
        [$col, $row] = $this->toIndexes($coordinate);
        $this->bufferCell($row, $col, $this->encodeExplicit($value, $dataType));

        return $this;
    }

    public function setCellValueExplicitByColumnAndRow(int $columnIndex, int $row, mixed $value, string $dataType = DataType::TYPE_STRING): static
    {
        $this->bufferCell($row, $columnIndex, $this->encodeExplicit($value, $dataType));

        return $this;
    }

    /**
     * Bulk load — the fast path. Chunks go straight to the extension without
     * touching the per-cell buffer.
     *
     * @param array<int|string, array<int|string, mixed>> $source
     */
    public function fromArray(array $source, mixed $nullValue = null, string $startCell = 'A1', bool $strictNullComparison = false): static
    {
        if (!\is_array(\reset($source))) {
            $source = [$source]; // single row, like PhpSpreadsheet
        }
        [$startCol, $startRow] = Coordinate::indexesFromString($startCell);
        $this->flush();

        $chunk = [];
        $chunkStart = $startRow;
        $row = $startRow;
        foreach ($source as $rowData) {
            $encoded = [];
            foreach ($rowData as $value) {
                $isNull = $strictNullComparison ? $value === $nullValue : $value == $nullValue;
                $encoded[] = $isNull ? null : $this->bindValue($value);
            }
            $chunk[] = $encoded;
            ++$row;
            if (\count($chunk) >= 1024) {
                Native::writeRows($this->parent->getHandle(), $this->title, $chunkStart, $startCol, $chunk);
                $chunk = [];
                $chunkStart = $row;
            }
        }
        if ($chunk !== []) {
            Native::writeRows($this->parent->getHandle(), $this->title, $chunkStart, $startCol, $chunk);
        }

        return $this;
    }

    // --- reads ----------------------------------------------------------------

    public function getCell(string|array $coordinate): Cell
    {
        [$col, $row] = $this->toIndexes($coordinate);

        return new Cell($this, Coordinate::stringFromColumnIndex($col) . $row);
    }

    public function getCellByColumnAndRow(int $columnIndex, int $row): Cell
    {
        return new Cell($this, Coordinate::stringFromColumnIndex($columnIndex) . $row);
    }

    /** @internal used by Cell */
    public function readCell(string $coordinate, int $mode): mixed
    {
        $this->flush();

        return Native::getCell($this->parent->getHandle(), $this->title, $coordinate, $mode);
    }

    public function toArray(mixed $nullValue = null, bool $calculateFormulas = true, bool $formatData = true, bool $returnCellRef = false): array
    {
        $this->flush();
        [$maxRow, $maxCol] = Native::dimensions($this->parent->getHandle(), $this->title);

        return $this->collectRows(1, \max($maxRow, 1), 1, \max($maxCol, 1), $nullValue, $formatData, $returnCellRef);
    }

    public function rangeToArray(string $range, mixed $nullValue = null, bool $calculateFormulas = true, bool $formatData = true, bool $returnCellRef = false): array
    {
        $this->flush();
        [[$startCol, $startRow], [$endCol, $endRow]] = Coordinate::rangeBoundaries($range);

        return $this->collectRows($startRow, $endRow, $startCol, $endCol, $nullValue, $formatData, $returnCellRef);
    }

    /**
     * Chunked extension reads (1k rows per CGO call, PLAN.md §6) assembled
     * into PhpSpreadsheet's toArray() shape.
     */
    private function collectRows(int $startRow, int $endRow, int $startCol, ?int $endCol, mixed $nullValue, bool $formatData, bool $returnCellRef): array
    {
        $handle = $this->parent->getHandle();
        $out = [];
        $row = $startRow;
        while ($row <= $endRow) {
            $want = \min(1000, $endRow - $row + 1);
            [$rows, $more] = Native::readRows($handle, $this->title, $row, $want, !$formatData);
            $got = \count($rows);
            foreach ($rows as $i => $cols) {
                $rowNum = $row + $i;
                $slice = [];
                $last = $endCol ?? \max(\count($cols), 1);
                for ($c = $startCol; $c <= $last; ++$c) {
                    $v = $cols[$c - 1] ?? null;
                    if ($v === null || $v === '') {
                        $v = $nullValue;
                    } elseif (!$formatData && \is_string($v) && \is_numeric($v) && !\str_starts_with($v, '=')) {
                        // raw mode returns native types like PhpSpreadsheet
                        $v = $v + 0;
                    }
                    if ($returnCellRef) {
                        $slice[Coordinate::stringFromColumnIndex($c)] = $v;
                    } else {
                        $slice[] = $v;
                    }
                }
                if ($returnCellRef) {
                    $out[$rowNum] = $slice;
                } else {
                    $out[] = $slice;
                }
            }
            $row += $got;
            if (!$more || $got === 0) {
                break;
            }
        }
        // pad trailing missing rows so dimensions stay faithful
        while ($row <= $endRow) {
            $width = ($endCol ?? 1) - $startCol + 1;
            $pad = \array_fill(0, \max($width, 1), $nullValue);
            if ($returnCellRef) {
                $out[$row] = \array_combine(
                    \array_map(Coordinate::stringFromColumnIndex(...), \range($startCol, $endCol ?? $startCol)),
                    $pad
                );
            } else {
                $out[] = $pad;
            }
            ++$row;
        }

        return $out;
    }

    public function getHighestRow(?string $column = null): int
    {
        $this->flush();
        [$maxRow] = Native::dimensions($this->parent->getHandle(), $this->title);

        return \max($maxRow, 1);
    }

    public function getHighestColumn(?int $row = null): string
    {
        $this->flush();
        [, $maxCol] = Native::dimensions($this->parent->getHandle(), $this->title);

        return Coordinate::stringFromColumnIndex(\max($maxCol, 1));
    }

    public function getHighestDataRow(?string $column = null): int
    {
        return $this->getHighestRow($column);
    }

    public function getHighestDataColumn(?int $row = null): string
    {
        return $this->getHighestColumn($row);
    }

    // --- structure / styling ----------------------------------------------------

    public function mergeCells(string|array $range): static
    {
        if (\is_array($range)) {
            $range = \implode(':', \array_map(
                static fn (array $c): string => Coordinate::stringFromColumnIndex($c[0]) . $c[1],
                $range
            ));
        }
        // no flush: the Go op-log applies merges by coordinates regardless of
        // when the rows arrive, and flushing here would end streaming early
        Native::mergeCells($this->parent->getHandle(), $this->title, $range);

        return $this;
    }

    public function getStyle(string|array $cellCoordinate): Style
    {
        if (\is_array($cellCoordinate)) {
            $cellCoordinate = Coordinate::stringFromColumnIndex($cellCoordinate[0]) . $cellCoordinate[1];
        }

        return new Style($this, $cellCoordinate);
    }

    public function getColumnDimension(string $column): ColumnDimension
    {
        return new ColumnDimension($this, $column);
    }

    public function getColumnDimensionByColumn(int $columnIndex): ColumnDimension
    {
        return new ColumnDimension($this, Coordinate::stringFromColumnIndex($columnIndex));
    }

    public function getRowDimension(int $row): RowDimension
    {
        return new RowDimension($this, $row);
    }

    public function setAutoFilter(string|array $range): static
    {
        if (\is_array($range)) {
            $range = \implode(':', \array_map(
                static fn (array $c): string => Coordinate::stringFromColumnIndex($c[0]) . $c[1],
                $range
            ));
        }
        Native::autoFilter($this->parent->getHandle(), $this->title, $range);

        return $this;
    }

    public function freezePane(?string $coordinate): static
    {
        Native::freezePanes($this->parent->getHandle(), $this->title, $coordinate ?? '');

        return $this;
    }

    public function freezePaneByColumnAndRow(int $columnIndex, int $row): static
    {
        return $this->freezePane(Coordinate::stringFromColumnIndex($columnIndex) . $row);
    }

    public function unfreezePane(): static
    {
        return $this->freezePane(null);
    }

    public function getComment(string|array $cellCoordinate): Comment
    {
        if (\is_array($cellCoordinate)) {
            $cellCoordinate = Coordinate::stringFromColumnIndex($cellCoordinate[0]) . $cellCoordinate[1];
        }

        return new Comment($this, $cellCoordinate);
    }

    public function getCommentByColumnAndRow(int $columnIndex, int $row): Comment
    {
        return $this->getComment(Coordinate::stringFromColumnIndex($columnIndex) . $row);
    }

    public function setHyperlink(string $cellCoordinate, ?Hyperlink $hyperlink): static
    {
        Native::setHyperlink(
            $this->parent->getHandle(),
            $this->title,
            $cellCoordinate,
            $hyperlink?->getUrl() ?? '',
            $hyperlink?->getTooltip() ?? '',
        );

        return $this;
    }

    public function getPageSetup(): PageSetup
    {
        return $this->pageSetup ??= new PageSetup($this);
    }

    // --- internals -----------------------------------------------------------------

    /** @internal flush the write-behind buffer to the extension */
    public function flush(): void
    {
        if ($this->buffer === []) {
            return;
        }
        $buffer = $this->buffer;
        $this->buffer = [];
        $this->bufferedRows = 0;
        \ksort($buffer);

        $handle = $this->parent->getHandle();
        $runRows = [];
        $runStart = null;
        $runMinCol = PHP_INT_MAX;
        $prevRow = null;

        $send = function () use (&$runRows, &$runStart, &$runMinCol, $handle): void {
            if ($runRows === []) {
                return;
            }
            $payload = [];
            foreach ($runRows as $cols) {
                $maxCol = \max(\array_keys($cols));
                $rowOut = \array_fill(0, $maxCol - $runMinCol + 1, null);
                foreach ($cols as $col => $value) {
                    $rowOut[$col - $runMinCol] = $value;
                }
                $payload[] = $rowOut;
            }
            Native::writeRows($handle, $this->title, $runStart, $runMinCol, $payload);
            $runRows = [];
            $runMinCol = PHP_INT_MAX;
        };

        foreach ($buffer as $row => $cols) {
            if ($prevRow !== null && $row !== $prevRow + 1) {
                $send();
                $runStart = null;
            }
            if ($runStart === null) {
                $runStart = $row;
            }
            $runMinCol = \min($runMinCol, \min(\array_keys($cols)));
            $runRows[] = $cols;
            $prevRow = $row;
        }
        $send();
    }

    private function bufferCell(int $row, int $col, mixed $encoded): void
    {
        if (!isset($this->buffer[$row])) {
            ++$this->bufferedRows;
        }
        $this->buffer[$row][$col] = $encoded;
        if ($this->bufferedRows >= self::FLUSH_AFTER_ROWS) {
            $this->flush();
        }
    }

    /** DefaultValueBinder parity for values the Go side cannot bind itself. */
    private function bindValue(mixed $value): mixed
    {
        if ($value instanceof \DateTimeInterface) {
            return [Native::MARK_NUMERIC, Date::dateTimeToExcel($value)];
        }
        if ($value instanceof \Stringable) {
            return (string) $value;
        }
        if ($value !== null && !\is_scalar($value)) {
            throw new Exception('Unsupported cell value of type ' . \get_debug_type($value));
        }

        return $value;
    }

    /** @return array{0: string, 1: mixed} */
    private function encodeExplicit(mixed $value, string $dataType): array
    {
        return match ($dataType) {
            DataType::TYPE_STRING, DataType::TYPE_STRING2, DataType::TYPE_INLINE => [Native::MARK_STRING, (string) $value],
            DataType::TYPE_NUMERIC => [Native::MARK_NUMERIC, $value],
            DataType::TYPE_BOOL => [Native::MARK_BOOL, $value],
            DataType::TYPE_FORMULA => [Native::MARK_FORMULA, (string) $value],
            DataType::TYPE_NULL => [Native::MARK_STRING, ''],
            default => throw new Exception("Invalid datatype: $dataType"),
        };
    }

    /** @return array{0: int, 1: int} [col, row] */
    private function toIndexes(string|array $coordinate): array
    {
        if (\is_array($coordinate)) {
            return [(int) $coordinate[0], (int) $coordinate[1]];
        }

        return Coordinate::indexesFromString($coordinate);
    }

    /** @internal called on rename from Spreadsheet bookkeeping */
    public function syncTitle(string $title): void
    {
        $this->title = $title;
    }
}
