<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet\PivotTable;

use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Shared\StringHelper;
use EasyExcel\Compat\Worksheet\Worksheet;

/**
 * Builds a new pivot table from a source range, PhpSpreadsheet 5.10 style:
 * field names come from the range's header row, fields are placed on the
 * row / column / page / value axes, and build() registers the table on the
 * target sheet. Writer\Xlsx writes it through excelize.AddPivotTable with
 * refresh-on-load set, so the values are computed when the file opens.
 *
 * The source range is read in bulk (rangeToArray, 1k rows per extension call)
 * rather than cell by cell.
 */
class PivotTableBuilder
{
    private readonly string $sourceRange;

    /** @var list<string> in source-column order */
    private readonly array $fieldNames;

    /** @var array<string, array{axis: string, subtotal: ?string, caption: ?string}> */
    private array $placements = [];

    /** @var array<string, PivotFieldGroup> */
    private array $groups = [];

    /** @param string $sourceRange the source range including its header row, e.g. "A1:C100" */
    public function __construct(private readonly Worksheet $sourceWorksheet, string $sourceRange)
    {
        $this->sourceRange = \str_replace('$', '', $sourceRange);
        $this->fieldNames = $this->readHeaderNames();
    }

    public function addRowField(string $fieldName): self
    {
        return $this->place($fieldName, PivotField::AXIS_ROW, null, null);
    }

    public function addColumnField(string $fieldName): self
    {
        return $this->place($fieldName, PivotField::AXIS_COLUMN, null, null);
    }

    /** Places a field on the report-filter axis. */
    public function addPageField(string $fieldName): self
    {
        return $this->place($fieldName, PivotField::AXIS_PAGE, null, null);
    }

    /** Stored on the cache definition but not written (COMPAT.md divergence 33). */
    public function groupFieldByNumericRange(string $fieldName, float $interval, ?float $startNum = null, ?float $endNum = null): self
    {
        $this->assertFieldExists($fieldName);
        $this->groups[$fieldName] = PivotFieldGroup::numeric($interval, $startNum, $endNum);

        return $this;
    }

    /**
     * Stored on the cache definition but not written (COMPAT.md divergence 33).
     *
     * @param string|array<string> $groupBy one or more PivotFieldGroup::GROUP_BY_* constants
     */
    public function groupFieldByDate(string $fieldName, array|string $groupBy, ?string $startDate = null, ?string $endDate = null): self
    {
        $this->assertFieldExists($fieldName);
        $this->groups[$fieldName] = PivotFieldGroup::date($groupBy, $startDate, $endDate);

        return $this;
    }

    /** @param string $subtotal one of the PivotField::SUBTOTAL_* constants */
    public function addDataField(string $fieldName, string $subtotal = PivotField::SUBTOTAL_SUM, ?string $caption = null): self
    {
        return $this->place($fieldName, PivotField::AXIS_VALUES, $subtotal, $caption);
    }

    /** @param string $targetCell top-left cell of the pivot table, e.g. "A3" */
    public function build(Worksheet $targetWorksheet, string $targetCell, string $name = 'PivotTable1'): PivotTable
    {
        $this->assertHasDataField();

        $cache = (new PivotCacheDefinition(1))
            ->setSourceWorksheet($this->sourceWorksheet->getTitle())
            ->setSourceRange($this->sourceRange)
            ->setCacheFields($this->fieldNames);
        foreach ($this->distinctValues() as $fieldName => $values) {
            // a numeric header ("2024") comes back as an int array key
            $fieldName = (string) $fieldName;
            $cache->setSharedItems($fieldName, $values);
            if (isset($this->groups[$fieldName])) {
                $cache->setFieldGroup($fieldName, $this->groups[$fieldName]);
            }
        }

        $targetCell = \str_replace('$', '', $targetCell);
        $pivotTable = (new PivotTable($name))
            ->setGenerated(true)
            ->setCacheDefinition($cache)
            // the extent is recomputed on refresh; a one-cell ref is the start
            ->setLocation("$targetCell:$targetCell");

        foreach ($this->fieldNames as $index => $fieldName) {
            $field = new PivotField($index, $fieldName);
            $placement = $this->placements[$fieldName] ?? null;
            if ($placement !== null && $placement['axis'] === PivotField::AXIS_VALUES) {
                $field->setDataField(true)
                    ->setSubtotal($placement['subtotal'])
                    ->setDataFieldCaption($placement['caption'] ?? self::defaultCaption($placement['subtotal'], $fieldName));
            } elseif ($placement !== null) {
                $field->setAxis($placement['axis']);
            }
            $pivotTable->addField($field);
        }
        $targetWorksheet->addPivotTable($pivotTable);

        return $pivotTable;
    }

    private function place(string $fieldName, string $axis, ?string $subtotal, ?string $caption): self
    {
        $this->assertFieldExists($fieldName);
        $this->placements[$fieldName] = ['axis' => $axis, 'subtotal' => $subtotal, 'caption' => $caption];

        return $this;
    }

    private function assertFieldExists(string $fieldName): void
    {
        if (!\in_array($fieldName, $this->fieldNames, true)) {
            throw new Exception("Pivot source field '{$fieldName}' does not exist in range {$this->sourceRange}");
        }
    }

    private function assertHasDataField(): void
    {
        foreach ($this->placements as $placement) {
            if ($placement['axis'] === PivotField::AXIS_VALUES) {
                return;
            }
        }

        throw new Exception('A pivot table requires at least one data (value) field');
    }

    /** @return list<string> */
    private function readHeaderNames(): array
    {
        [[$startCol, $headerRow], [$endCol]] = Coordinate::rangeBoundaries($this->sourceRange);

        return \array_map(
            static fn (mixed $v): string => StringHelper::convertToString($v, false),
            // formatted text: it is what excelize resolves field names against,
            // and a raw read would turn a text header like "007" into 7
            $this->readRows($startCol, $headerRow, $endCol, $headerRow, true)[0] ?? [],
        );
    }

    /** @return array<string, list<string>> distinct non-empty values per field, first-seen order */
    private function distinctValues(): array
    {
        [[$startCol, $headerRow], [$endCol, $lastRow]] = Coordinate::rangeBoundaries($this->sourceRange);
        $seen = \array_fill_keys($this->fieldNames, []);
        if ($lastRow > $headerRow) {
            foreach ($this->readRows($startCol, $headerRow + 1, $endCol, $lastRow) as $row) {
                foreach ($this->fieldNames as $i => $fieldName) {
                    $value = StringHelper::convertToString($row[$i] ?? null, false);
                    if ($value !== '') {
                        $seen[$fieldName][$value] = true;
                    }
                }
            }
        }

        return \array_map(static fn (array $values): array => \array_map('strval', \array_keys($values)), $seen);
    }

    /** @return list<list<mixed>> raw values (formulas as source), like Cell::getValue() */
    private function readRows(int $startCol, int $startRow, int $endCol, int $endRow, bool $formatted = false): array
    {
        $range = Coordinate::stringFromColumnIndex($startCol) . $startRow
            . ':' . Coordinate::stringFromColumnIndex($endCol) . $endRow;

        return \array_map(
            'array_values',
            \array_values($this->sourceWorksheet->rangeToArray($range, null, false, $formatted)),
        );
    }

    private static function defaultCaption(?string $subtotal, string $fieldName): string
    {
        $label = match ($subtotal) {
            PivotField::SUBTOTAL_COUNT, PivotField::SUBTOTAL_COUNT_NUMS => 'Count',
            PivotField::SUBTOTAL_AVERAGE => 'Average',
            PivotField::SUBTOTAL_MAX => 'Max',
            PivotField::SUBTOTAL_MIN => 'Min',
            PivotField::SUBTOTAL_PRODUCT => 'Product',
            PivotField::SUBTOTAL_STD_DEV => 'StdDev',
            PivotField::SUBTOTAL_STD_DEV_P => 'StdDevp',
            PivotField::SUBTOTAL_VAR => 'Var',
            PivotField::SUBTOTAL_VAR_P => 'Varp',
            default => 'Sum',
        };

        return "{$label} of {$fieldName}";
    }
}
