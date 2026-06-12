<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Native;

/** Row height facade for one row. */
class RowDimension
{
    private float $height = -1;

    public function __construct(private Worksheet $worksheet, private int $rowIndex)
    {
    }

    public function getRowIndex(): int
    {
        return $this->rowIndex;
    }

    public function setRowHeight(float|int $height): static
    {
        $this->height = (float) $height;
        Native::setRowHeight(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->rowIndex,
            (float) $height,
        );

        return $this;
    }

    public function getRowHeight(): float
    {
        return $this->height;
    }
}
