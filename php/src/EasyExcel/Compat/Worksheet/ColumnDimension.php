<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Native;

/** Column width / auto-size facade for one column. */
class ColumnDimension
{
    private float $width = -1;
    private bool $autoSize = false;

    public function __construct(private Worksheet $worksheet, private string $columnIndex)
    {
    }

    public function getColumnIndex(): string
    {
        return $this->columnIndex;
    }

    public function setWidth(float|int $width): static
    {
        $this->width = (float) $width;
        $col = Coordinate::columnIndexFromString($this->columnIndex);
        Native::setColWidth(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $col,
            $col,
            (float) $width,
        );

        return $this;
    }

    public function getWidth(): float
    {
        return $this->width;
    }

    /** Auto width is approximated at save time by character count (COMPAT.md). */
    public function setAutoSize(bool $autoSize): static
    {
        $this->autoSize = $autoSize;
        if ($autoSize) {
            $col = Coordinate::columnIndexFromString($this->columnIndex);
            Native::setColAutoSize(
                $this->worksheet->getParent()->getHandle(),
                $this->worksheet->getTitle(),
                $col,
                $col,
            );
        }

        return $this;
    }

    public function getAutoSize(): bool
    {
        return $this->autoSize;
    }
}
