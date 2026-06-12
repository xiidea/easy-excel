<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/** Border collection: top/bottom/left/right plus the allBorders pseudo-side. */
class Borders
{
    private array $sides = [];

    public function __construct(private Style $style)
    {
    }

    public function getTop(): Border
    {
        return $this->side('top');
    }

    public function getBottom(): Border
    {
        return $this->side('bottom');
    }

    public function getLeft(): Border
    {
        return $this->side('left');
    }

    public function getRight(): Border
    {
        return $this->side('right');
    }

    public function getAllBorders(): Border
    {
        return $this->side('allBorders');
    }

    public function getOutline(): Border
    {
        return $this->side('outline');
    }

    private function side(string $position): Border
    {
        return $this->sides[$position] ??= new Border($this->style, $position);
    }
}
