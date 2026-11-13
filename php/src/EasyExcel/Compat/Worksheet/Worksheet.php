<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Cell\Cell;
use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Compat\Cell\DataType;
use EasyExcel\Compat\Cell\Hyperlink;
use EasyExcel\Compat\Cell\IValueBinder;
use EasyExcel\Compat\Comment;
use EasyExcel\Compat\RichText\RichText;
use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Reader\Xlsx\PivotTableReader;
use EasyExcel\Compat\Reader\Xlsx\Sparklines as SparklinesReader;
use EasyExcel\Compat\Shared\Date;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Style\Color;
use EasyExcel\Compat\Style\Style;
use EasyExcel\Compat\Worksheet\PivotTable\PivotField;
use EasyExcel\Compat\Worksheet\PivotTable\PivotTable;
use EasyExcel\Compat\Worksheet\Sparkline\Sparkline;
use EasyExcel\Compat\Worksheet\Sparkline\SparklineGroup;
use EasyExcel\Compat\Worksheet\Sparkline\SparklineType;
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

    public const BREAK_NONE = 0;

    public const BREAK_ROW = 1;

    public const BREAK_COLUMN = 2;

    public const BREAK_ROW_MAX_COLUMN = 16383;

    /** @var array<int, array<int, mixed>> row => col => scalar|[marker, value] */
    private array $buffer = [];

    private int $bufferedRows = 0;

    private ?PageSetup $pageSetup = null;

    private ?Protection $protection = null;

    /** @var list<Drawing> List of attached drawings */
    private array $drawings = [];

    /**
     * PhpSpreadsheet parity: a Worksheet may be constructed detached
     * (`new Worksheet()`) and attached later via Spreadsheet::addSheet().
     * Until attached, only title get/set work; anything touching the native
     * workbook throws (see workbookHandle()).
     */
    public function __construct(private ?Spreadsheet $parent = null, private string $title = 'Worksheet')
    {
    }

    public function getParent(): ?Spreadsheet
    {
        return $this->parent;
    }

    /** @internal called by Spreadsheet::addSheet() when attaching a detached sheet */
    public function rebindParent(Spreadsheet $parent): static
    {
        if ($this->parent === $parent) {
            return $this;
        }
        if ($this->parent !== null) {
            throw new Exception(
                'Worksheet "' . $this->title . '" already belongs to a workbook; moving sheets between workbooks is not supported'
            );
        }
        $this->parent = $parent;

        return $this;
    }

    private function workbookHandle(): int
    {
        if ($this->parent === null) {
            throw new Exception(
                'Worksheet "' . $this->title . '" is not attached to a Spreadsheet yet — add it with $spreadsheet->addSheet() first'
            );
        }

        return $this->parent->getHandle();
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
        if ($this->parent !== null) {
            Native::renameSheet($this->parent->getHandle(), $this->title, $title);
        }
        $this->title = $title;

        return $this;
    }

    // --- drawings -------------------------------------------------------------

    public function addDrawing(Drawing $drawing): void
    {
        $this->drawings[] = $drawing;
    }

    /** @return list<Drawing> */
    public function getDrawingCollection(): array
    {
        return $this->drawings;
    }

    // --- writes ---------------------------------------------------------------

    public function setCellValue(string|array $coordinate, mixed $value): static
    {
        if ($value instanceof RichText) {
            [$col, $row] = $this->toIndexes($coordinate);
            $this->setRichTextValue(Coordinate::stringFromColumnIndex($col) . $row, $value);

            return $this;
        }
        if (($binder = $this->activeBinder()) !== null) {
            [$col, $row] = $this->toIndexes($coordinate);
            $binder->bindValue(new Cell($this, Coordinate::stringFromColumnIndex($col) . $row), $value);

            return $this;
        }
        [$col, $row] = $this->toIndexes($coordinate);
        $this->bufferCell($row, $col, $this->bindValue($value));

        return $this;
    }

    /**
     * Effective custom binder: the workbook-level binder
     * (Spreadsheet::setValueBinder(), PhpSpreadsheet >= 2.x) wins over the
     * legacy process-wide Cell::setValueBinder(); null means default binding.
     */
    private function activeBinder(): ?IValueBinder
    {
        return $this->parent?->getValueBinder() ?? Cell::customValueBinder();
    }

    /** @internal rich-text cell value (wave 4.4): queued via the extension */
    public function setRichTextValue(string $coordinate, RichText $richText): void
    {
        // a plain placeholder keeps dimensions correct; the extension applies
        // the formatted runs over it at save (COMPAT.md)
        [$col, $row] = Coordinate::indexesFromString($coordinate);
        $this->bufferCell($row, $col, $richText->getPlainText());
        $this->flush();
        Native::setRichText($this->workbookHandle(), $this->title, $coordinate, $richText->toRunSpecs());
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

        // a custom binder must see every value: route through the buffered
        // per-cell path (slower, still batched per 512 rows)
        if (($binder = $this->activeBinder()) !== null) {
            $row = $startRow;
            foreach ($source as $rowData) {
                $col = $startCol;
                foreach ($rowData as $value) {
                    $isNull = $strictNullComparison ? $value === $nullValue : $value == $nullValue;
                    if (!$isNull) {
                        $binder->bindValue(new Cell($this, Coordinate::stringFromColumnIndex($col) . $row), $value);
                    }
                    ++$col;
                }
                ++$row;
            }

            return $this;
        }
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
                Native::writeRows($this->workbookHandle(), $this->title, $chunkStart, $startCol, $chunk);
                $chunk = [];
                $chunkStart = $row;
            }
        }
        if ($chunk !== []) {
            Native::writeRows($this->workbookHandle(), $this->title, $chunkStart, $startCol, $chunk);
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
        if (($filter = $this->parent->getReadFilter()) !== null) {
            [$col, $row] = Coordinate::indexesFromString($coordinate);
            if (!$filter->readCell(Coordinate::stringFromColumnIndex($col), $row, $this->title)) {
                return null;
            }
        }

        return Native::getCell($this->workbookHandle(), $this->title, $coordinate, $mode);
    }

    public function toArray(mixed $nullValue = null, bool $calculateFormulas = true, bool $formatData = true, bool $returnCellRef = false): array
    {
        $this->flush();
        [$maxRow, $maxCol] = Native::dimensions($this->workbookHandle(), $this->title);

        return $this->collectRows(1, \max($maxRow, 1), 1, \max($maxCol, 1), $nullValue, $formatData, $returnCellRef, $calculateFormulas);
    }

    public function rangeToArray(string $range, mixed $nullValue = null, bool $calculateFormulas = true, bool $formatData = true, bool $returnCellRef = false): array
    {
        $this->flush();
        [[$startCol, $startRow], [$endCol, $endRow]] = Coordinate::rangeBoundaries($range);

        return $this->collectRows($startRow, $endRow, $startCol, $endCol, $nullValue, $formatData, $returnCellRef, $calculateFormulas);
    }

    /**
     * Chunked extension reads (1k rows per CGO call, PLAN.md §6) assembled
     * into PhpSpreadsheet's toArray() shape.
     */
    private function collectRows(int $startRow, int $endRow, int $startCol, ?int $endCol, mixed $nullValue, bool $formatData, bool $returnCellRef, bool $calc = false): array
    {
        $handle = $this->workbookHandle();
        $filter = $this->parent->getReadFilter();
        $out = [];
        $row = $startRow;
        while ($row <= $endRow) {
            $want = \min(1000, $endRow - $row + 1);
            [$rows, $more] = Native::readRows($handle, $this->title, $row, $want, !$formatData, $calc);
            $got = \count($rows);
            foreach ($rows as $i => $cols) {
                $rowNum = $row + $i;
                $slice = [];
                $last = $endCol ?? \max(\count($cols), 1);
                for ($c = $startCol; $c <= $last; ++$c) {
                    $v = $cols[$c - 1] ?? null;
                    if ($filter !== null && $v !== null && $v !== ''
                        && !$filter->readCell(Coordinate::stringFromColumnIndex($c), $rowNum, $this->title)) {
                        $v = null;
                    }
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
        [$maxRow] = Native::dimensions($this->workbookHandle(), $this->title);

        return \max($maxRow, 1);
    }

    public function getHighestColumn(?int $row = null): string
    {
        $this->flush();
        [, $maxCol] = Native::dimensions($this->workbookHandle(), $this->title);

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
        Native::mergeCells($this->workbookHandle(), $this->title, $range);

        return $this;
    }

    public function unmergeCells(string|array $range): static
    {
        if (\is_array($range)) {
            $range = \implode(':', \array_map(
                static fn (array $c): string => Coordinate::stringFromColumnIndex($c[0]) . $c[1],
                $range
            ));
        }
        Native::unmergeCells($this->workbookHandle(), $this->title, $range);

        return $this;
    }

    /** @return array<string, string> range => range, PhpSpreadsheet shape */
    public function getMergeCells(): array
    {
        $this->flush();
        $out = [];
        foreach (Native::getMerges($this->workbookHandle(), $this->title) as $range) {
            $out[$range] = $range;
        }

        return $out;
    }

    public function getStyle(string|array $cellCoordinate): Style
    {
        if (\is_array($cellCoordinate)) {
            $cellCoordinate = Coordinate::stringFromColumnIndex($cellCoordinate[0]) . $cellCoordinate[1];
        }

        return new Style($this, $cellCoordinate);
    }

    /** @var array<string, list<\EasyExcel\Compat\Style\Conditional>> */
    private array $conditionalRegistry = [];

    /** @internal rules set this session, keyed by range (Style::getConditionalStyles) */
    public function rememberConditionalStyles(string $range, array $rules): void
    {
        $this->conditionalRegistry[$range] = $rules;
    }

    /** @internal */
    public function recallConditionalStyles(string $range): array
    {
        return $this->conditionalRegistry[$range] ?? [];
    }

    /** @var array<string, ColumnDimension> */
    private array $columnDimensions = [];

    /**
     * Cap on retained RowDimension objects; see getRowDimension(). Chosen to
     * cover realistic "style the first N rows" patterns without letting a
     * per-row loop over a large export grow without bound.
     */
    private const ROW_DIMENSION_CACHE_LIMIT = 4096;

    /** @var array<int, RowDimension> */
    private array $rowDimensions = [];

    /**
     * Dimensions are cached per column/row so a width or height written
     * through one call is readable through the next. PhpSpreadsheet's Html
     * writer reads the width back via getWidth(), so handing out a fresh
     * instance each call made every column fall back to the writer's default.
     */
    public function getColumnDimension(string $column): ColumnDimension
    {
        // Normalised: column refs are case-insensitive in Excel, so 'a' and
        // 'A' must share one dimension. Keying on the raw string let two
        // callers set different widths on the same column, last save winning
        // (review item 5).
        $column = \strtoupper($column);

        return $this->columnDimensions[$column] ??= new ColumnDimension($this, $column);
    }

    public function getColumnDimensionByColumn(int $columnIndex): ColumnDimension
    {
        return $this->getColumnDimension(Coordinate::stringFromColumnIndex($columnIndex));
    }

    /**
     * Row dimensions are cached like column dimensions, but rows are unbounded
     * where columns are not: a per-row getRowDimension()->setRowHeight() loop
     * over a 500k-row export would retain 500k objects for the sheet's
     * lifetime. The cache is capped; past the cap a fresh instance is handed
     * out, which still writes through correctly and only loses read-back of a
     * height set earlier in the same session (review item 5).
     */
    public function getRowDimension(int $row): RowDimension
    {
        if (isset($this->rowDimensions[$row])) {
            return $this->rowDimensions[$row];
        }
        if (\count($this->rowDimensions) >= self::ROW_DIMENSION_CACHE_LIMIT) {
            return new RowDimension($this, $row);
        }

        return $this->rowDimensions[$row] = new RowDimension($this, $row);
    }

    private ?RowDimension $defaultRowDimension = null;

    private ?ColumnDimension $defaultColumnDimension = null;

    /**
     * Sheet-wide default row height (sheetFormatPr). Unlike per-row heights,
     * this stays free on streaming sheets — prefer it for uniform heights.
     */
    public function getDefaultRowDimension(): RowDimension
    {
        return $this->defaultRowDimension ??= new RowDimension($this, null);
    }

    /** Sheet-wide default column width (sheetFormatPr). */
    public function getDefaultColumnDimension(): ColumnDimension
    {
        return $this->defaultColumnDimension ??= new ColumnDimension($this, null);
    }

    private string $autoFilterRange = '';

    private ?AutoFilter $autoFilter = null;

    public function setAutoFilter(string|array $range): static
    {
        if (\is_array($range)) {
            $range = \implode(':', \array_map(
                static fn (array $c): string => Coordinate::stringFromColumnIndex($c[0]) . $c[1],
                $range
            ));
        }
        $this->autoFilterRange = $range;
        Native::autoFilter($this->workbookHandle(), $this->title, $range);

        return $this;
    }

    public function getAutoFilter(): AutoFilter
    {
        return $this->autoFilter ??= new AutoFilter($this);
    }

    /** @internal the range set this session (loaded files are not introspected) */
    public function autoFilterRange(): string
    {
        return $this->autoFilterRange;
    }

    public function getRowIterator(int $startRow = 1, ?int $endRow = null): RowIterator
    {
        $this->flush();

        return new RowIterator($this, $startRow, $endRow);
    }

    public function getColumnIterator(string $startColumn = 'A', ?string $endColumn = null): ColumnIterator
    {
        $this->flush();

        return new ColumnIterator($this, $startColumn, $endColumn);
    }

    /** Copies a style (its native state for attached styles) onto a range. */
    public function duplicateStyle(Style $style, string $range): static
    {
        $spec = $style->describe();
        if ($spec !== []) {
            Native::applyStyle($this->workbookHandle(), $this->title, $range, $spec);
        }

        return $this;
    }

    public function freezePane(?string $coordinate): static
    {
        Native::freezePanes($this->workbookHandle(), $this->title, $coordinate ?? '');

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

    /**
     * Set a manual page break at a cell (wave 5.3).
     *
     * `BREAK_ROW` splits above the cell's row, `BREAK_COLUMN` to the left of
     * its column, and `BREAK_NONE` removes a break there. Applied at save like
     * the rest of the non-streamable surface (COMPAT.md divergence 11).
     */
    public function setBreak(string|array $coordinate, int $break = self::BREAK_NONE): static
    {
        if (\is_array($coordinate)) {
            $coordinate = Coordinate::stringFromColumnIndex($coordinate[0]) . $coordinate[1];
        }
        Native::setBreak($this->workbookHandle(), $this->title, $coordinate, $break);

        return $this;
    }

    public function setBreakByColumnAndRow(int $columnIndex, int $row, int $break = self::BREAK_NONE): static
    {
        return $this->setBreak(Coordinate::stringFromColumnIndex($columnIndex) . $row, $break);
    }

    /**
     * Set the selected cell or range (wave 5.3).
     *
     * Merges into the sheet's existing pane state, so selecting a cell after
     * freezePane() keeps the freeze rather than silently undoing it.
     */
    public function setSelectedCells(string|array $coordinate): static
    {
        if (\is_array($coordinate)) {
            $coordinate = Coordinate::stringFromColumnIndex($coordinate[0]) . $coordinate[1];
        }
        Native::setSelection($this->workbookHandle(), $this->title, $coordinate);

        return $this;
    }

    public function setSelectedCell(string $coordinate): static
    {
        return $this->setSelectedCells($coordinate);
    }

    public function setSelectedCellByColumnAndRow(int $columnIndex, int $row): static
    {
        return $this->setSelectedCells(Coordinate::stringFromColumnIndex($columnIndex) . $row);
    }

    /**
     * Accepted no-op, returning $this for chaining.
     *
     * PhpSpreadsheet measures rendered text here to size auto-width columns.
     * easy-excel approximates auto-size in Go at save time instead (COMPAT.md
     * divergence 10), so there is nothing to precompute — but callers that
     * invoke this before saving should not have to branch on the engine.
     */
    public function calculateColumnWidths(): static
    {
        return $this;
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
            $this->workbookHandle(),
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

    public function getProtection(): Protection
    {
        return $this->protection ??= new Protection($this);
    }

    // --- structure editing (wave 4.3) ----------------------------------------

    public function insertNewRowBefore(int $before, int $count = 1): static
    {
        $this->flush();
        Native::insertRows($this->workbookHandle(), $this->title, $before, $count);

        return $this;
    }

    public function removeRow(int $row, int $count = 1): static
    {
        $this->flush();
        Native::removeRows($this->workbookHandle(), $this->title, $row, $count);

        return $this;
    }

    public function insertNewColumnBefore(string $before, int $count = 1): static
    {
        $this->flush();
        Native::insertCols($this->workbookHandle(), $this->title, Coordinate::columnIndexFromString($before), $count);

        return $this;
    }

    public function insertNewColumnBeforeByIndex(int $beforeColumnIndex, int $count = 1): static
    {
        $this->flush();
        Native::insertCols($this->workbookHandle(), $this->title, $beforeColumnIndex, $count);

        return $this;
    }

    public function removeColumn(string $column, int $count = 1): static
    {
        $this->flush();
        Native::removeCols($this->workbookHandle(), $this->title, Coordinate::columnIndexFromString($column), $count);

        return $this;
    }

    public function removeColumnByIndex(int $columnIndex, int $count = 1): static
    {
        $this->flush();
        Native::removeCols($this->workbookHandle(), $this->title, $columnIndex, $count);

        return $this;
    }

    // --- sheet view, headers/footers, margins (wave 4.3) -----------------------

    private ?SheetView $sheetView = null;

    public function getSheetView(): SheetView
    {
        return $this->sheetView ??= new SheetView($this);
    }

    public function setShowGridlines(bool $show): static
    {
        $this->getSheetView()->pushExtra('showGridlines', $show);

        return $this;
    }

    private ?Color $tabColor = null;

    public function getTabColor(): Color
    {
        return ($this->tabColor ??= new Color())->bind(
            fn (string $argb) => $this->getSheetView()->pushExtra('tabColor', \substr($argb, 2))
        );
    }

    private ?HeaderFooter $headerFooter = null;

    public function getHeaderFooter(): HeaderFooter
    {
        return $this->headerFooter ??= new HeaderFooter($this);
    }

    private ?PageMargins $pageMargins = null;

    public function getPageMargins(): PageMargins
    {
        return $this->pageMargins ??= new PageMargins($this);
    }

    public function setDataValidation(string $range, ?\EasyExcel\Compat\Cell\DataValidation $dataValidation): static
    {
        if ($dataValidation !== null) {
            Native::setValidation(
                $this->workbookHandle(),
                $this->title,
                $range,
                $dataValidation->toSpec(),
            );
        }

        return $this;
    }

    /**
     * easy-excel native chart API: a declarative spec straight to the
     * extension (the PhpSpreadsheet Chart facade builds on this).
     *
     * @param array<string, mixed> $spec ['type','series','title','legend','width','height']
     */
    public function addNativeChart(string $cell, array $spec): static
    {
        Native::addChart($this->workbookHandle(), $this->title, $cell, $spec);

        return $this;
    }

    /** PhpSpreadsheet-compatible chart (Chart\Chart, mapped to the native spec). */
    public function addChart(\EasyExcel\Compat\Chart\Chart $chart): static
    {
        Native::addChart(
            $this->workbookHandle(),
            $this->title,
            $chart->getTopLeftCell(),
            $chart->buildSpec(),
        );

        return $this;
    }

    /** @var \ArrayObject<int, SparklineGroup>|null */
    private ?\ArrayObject $sparklineGroups = null;

    /**
     * On a loaded workbook the first call also reads the groups the source
     * file carried (Reader\Xlsx\Sparklines over the extension's copy of the
     * sheet's sparkline <ext>). Every add goes through here, so the file's
     * groups are always in place before any new one. They are already in the
     * file, so they are marked written rather than sent again.
     *
     * @return \ArrayObject<int, SparklineGroup>
     */
    public function getSparklineGroupCollection(): \ArrayObject
    {
        if ($this->sparklineGroups !== null) {
            return $this->sparklineGroups;
        }
        $this->sparklineGroups = new \ArrayObject();
        $ext = $this->parent !== null ? Native::loadedSparklines($this->workbookHandle(), $this->title) : '';
        if ($ext !== '') {
            $xml = \simplexml_load_string(
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                . ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                . ' xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
                . ' xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
                . ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                . '<extLst>' . $ext . '</extLst></worksheet>',
            );
            try {
                if ($xml === false) {
                    throw new Exception("Could not parse the sparklines of worksheet \"{$this->title}\"");
                }
                (new SparklinesReader($this, $xml))->load();
            } catch (\Throwable $e) {
                // not half-loaded: groups left unfrozen would be re-sent on save
                $this->sparklineGroups = null;

                throw $e;
            }
            foreach ($this->sparklineGroups as $group) {
                $group->markWritten();
            }
        }

        return $this->sparklineGroups;
    }

    public function addSparklineGroup(SparklineGroup $sparklineGroup): static
    {
        $this->getSparklineGroupCollection()->append($sparklineGroup);

        return $this;
    }

    /** Wraps one sparkline in its own default-formatted group and returns the group. */
    public function addSparkline(Sparkline $sparkline, SparklineType $type = SparklineType::Line): SparklineGroup
    {
        $group = (new SparklineGroup())->setType($type)->addSparkline($sparkline);
        $this->addSparklineGroup($group);

        return $group;
    }

    /**
     * Groups already written by a save cannot be taken back out of the file
     * (excelize has no sparkline removal), so clearing them throws instead of
     * pretending to succeed.
     */
    public function removeSparklineGroupCollection(): static
    {
        foreach ($this->getSparklineGroupCollection() as $group) {
            if ($group->isWritten()) {
                throw new Exception('Sparkline groups already written by a save cannot be removed');
            }
        }
        $this->sparklineGroups = new \ArrayObject();

        return $this;
    }

    /**
     * @internal send groups not yet written to the extension (Writer\Xlsx save).
     * Empty groups are skipped, as PhpSpreadsheet writes nothing for them.
     */
    public function writeSparklines(): void
    {
        if ($this->sparklineGroups === null) {
            return; // never touched: nothing new, and the file's own groups ride along
        }
        foreach ($this->sparklineGroups as $group) {
            if ($group->isWritten() || $group->getSparklines() === []) {
                continue;
            }
            Native::addSparklineGroup($this->workbookHandle(), $this->title, $group->toSpec());
            $group->markWritten();
        }
    }

    /** @var \ArrayObject<int, PivotTable>|null */
    private ?\ArrayObject $pivotTables = null;

    /**
     * On a loaded workbook the first call also reads the pivot tables the
     * source file carried (Reader\Xlsx\PivotTableReader over the extension's
     * copy of their parts). Every add goes through here, so they are in place
     * before any new one.
     *
     * @return \ArrayObject<int, PivotTable>
     */
    public function getPivotTableCollection(): \ArrayObject
    {
        if ($this->pivotTables !== null) {
            return $this->pivotTables;
        }
        $this->pivotTables = new \ArrayObject();
        $loaded = $this->parent !== null ? Native::loadedPivotTables($this->workbookHandle(), $this->title) : [];
        foreach ($loaded as $parts) {
            $table = self::loadPivotXml($parts['table'] ?? '');
            $cache = ($parts['cache'] ?? '') !== '' ? self::loadPivotXml($parts['cache']) : null;
            (new PivotTableReader($this, $table, $cache))->load();
        }

        return $this->pivotTables;
    }

    private static function loadPivotXml(string $xml): \SimpleXMLElement
    {
        $element = \simplexml_load_string($xml, \SimpleXMLElement::class, 0, 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        if ($element === false) {
            throw new Exception('Could not parse a pivot table part of the source file');
        }

        return $element;
    }

    /** @return \ArrayObject<int, PivotTable> alias of getPivotTableCollection() */
    public function getPivotTables(): \ArrayObject
    {
        return $this->getPivotTableCollection();
    }

    public function addPivotTable(PivotTable $pivotTable): static
    {
        $pivotTable->setWorksheet($this);
        $this->getPivotTableCollection()->append($pivotTable);

        return $this;
    }

    /** @return list<string> */
    public function getPivotTableNames(): array
    {
        $names = [];
        foreach ($this->getPivotTableCollection() as $pivotTable) {
            $names[] = $pivotTable->getName();
        }

        return $names;
    }

    /** Case-insensitive, like PhpSpreadsheet. */
    public function getPivotTableByName(string $name): ?PivotTable
    {
        $name = \mb_strtoupper($name);
        foreach ($this->getPivotTableCollection() as $pivotTable) {
            if (\mb_strtoupper($pivotTable->getName()) === $name) {
                return $pivotTable;
            }
        }

        return null;
    }

    public function removePivotTableCollection(): static
    {
        $this->pivotTables = new \ArrayObject();

        return $this;
    }

    /** True once a save wrote generated pivot tables here (they must be replaced, not appended). */
    private bool $pivotTablesWritten = false;

    /**
     * @internal send the generated pivot tables to the extension (Writer\Xlsx
     * save). PhpSpreadsheet regenerates them on every save, so the whole set
     * is sent and replaces what an earlier save wrote; a sheet that never had
     * any skips the call. Loaded (non-generated) pivot tables are left alone.
     */
    public function writePivotTables(): void
    {
        $specs = [];
        // never touched: nothing generated, and the file's own tables ride along
        foreach ($this->pivotTables ?? [] as $pivotTable) {
            if ($pivotTable->isGenerated()) {
                $specs[] = self::pivotTableSpec($pivotTable);
            }
        }
        if ($specs === [] && !$this->pivotTablesWritten) {
            return;
        }
        // set first: a call that fails may still have written part of the set
        $this->pivotTablesWritten = true;
        Native::setPivotTables($this->workbookHandle(), $this->title, $specs);
        $this->pivotTablesWritten = $specs !== [];
    }

    /** @return array<string, mixed> the extension's pivot spec (extension/compat/pivot.go) */
    private static function pivotTableSpec(PivotTable $pivotTable): array
    {
        $cache = $pivotTable->getCacheDefinition();
        if ($cache === null || $cache->getSourceWorksheet() === null || $cache->getSourceRange() === null) {
            throw new Exception("Pivot table '{$pivotTable->getName()}' has no source cache definition");
        }
        $axis = static fn (array $fields): array => \array_map(
            static fn (PivotField $f): array => ['field' => $f->getName()],
            $fields,
        );

        return [
            'name' => $pivotTable->getName(),
            'sourceSheet' => $cache->getSourceWorksheet(),
            'sourceRange' => $cache->getSourceRange(),
            'location' => $pivotTable->getLocation(),
            'rows' => $axis($pivotTable->getRowFields()),
            'columns' => $axis($pivotTable->getColumnFields()),
            'filters' => $axis($pivotTable->getPageFields()),
            'data' => \array_map(
                static fn (PivotField $f): array => [
                    'field' => $f->getName(),
                    'subtotal' => $f->getSubtotal() ?? PivotField::SUBTOTAL_SUM,
                    'caption' => $f->getDataFieldCaption() ?? '',
                ],
                $pivotTable->getDataFields(),
            ),
        ];
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

        $handle = $this->workbookHandle();
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
