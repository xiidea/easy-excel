<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Worksheet\AutoFilter\Column;
use EasyExcel\Native;

/**
 * Auto-filter facade. Without column rules the range rides the streaming
 * container patch; adding a column rule (getColumn) switches to the model
 * path so the FilterColumn XML is written (wave 4.4).
 */
class AutoFilter
{
    /** @var array<string, Column> column letter => Column */
    private array $columns = [];

    public function __construct(private Worksheet $worksheet)
    {
    }

    public function getRange(): string
    {
        return $this->worksheet->autoFilterRange();
    }

    public function setRange(string $range): static
    {
        $this->worksheet->setAutoFilter($range);

        return $this;
    }

    public function getColumn(string $columnIndex): Column
    {
        return $this->columns[$columnIndex] ??= new Column($this, $columnIndex);
    }

    public function getColumnByOffset(int $offset): Column
    {
        [$rangeStart] = \explode(':', $this->getRange() ?: 'A1');
        $startCol = \EasyExcel\Compat\Cell\Coordinate::columnIndexFromString(
            \preg_replace('/\d+/', '', $rangeStart)
        );

        return $this->getColumn(\EasyExcel\Compat\Cell\Coordinate::stringFromColumnIndex($startCol + $offset));
    }

    /** @internal re-sends the full column-rule set to the extension */
    public function applyColumns(): void
    {
        $range = $this->getRange();
        if ($range === '') {
            return;
        }
        $cols = [];
        foreach ($this->columns as $letter => $column) {
            $expr = $column->toExpression();
            if ($expr !== '') {
                $cols[] = ['column' => $letter, 'expression' => $expr];
            }
        }
        if ($cols === []) {
            return;
        }
        Native::autoFilterColumns(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $range,
            $cols,
        );
    }
}
