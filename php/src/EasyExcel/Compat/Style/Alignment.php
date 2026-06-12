<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/** Alignment component. */
class Alignment
{
    public const HORIZONTAL_GENERAL = 'general';
    public const HORIZONTAL_LEFT = 'left';
    public const HORIZONTAL_RIGHT = 'right';
    public const HORIZONTAL_CENTER = 'center';
    public const HORIZONTAL_CENTER_CONTINUOUS = 'centerContinuous';
    public const HORIZONTAL_JUSTIFY = 'justify';
    public const HORIZONTAL_FILL = 'fill';
    public const HORIZONTAL_DISTRIBUTED = 'distributed';

    public const VERTICAL_BOTTOM = 'bottom';
    public const VERTICAL_TOP = 'top';
    public const VERTICAL_CENTER = 'center';
    public const VERTICAL_JUSTIFY = 'justify';
    public const VERTICAL_DISTRIBUTED = 'distributed';

    private string $horizontal = self::HORIZONTAL_GENERAL;
    private string $vertical = self::VERTICAL_BOTTOM;
    private bool $wrapText = false;
    private bool $shrinkToFit = false;
    private int $textRotation = 0;
    private int $indent = 0;

    public function __construct(private Style $style)
    {
    }

    public function setHorizontal(string $horizontal): static
    {
        $this->horizontal = $horizontal;
        $this->style->mergeComponent('alignment', ['horizontal' => $horizontal]);

        return $this;
    }

    public function setVertical(string $vertical): static
    {
        $this->vertical = $vertical;
        $this->style->mergeComponent('alignment', ['vertical' => $vertical]);

        return $this;
    }

    public function setWrapText(bool $wrapText): static
    {
        $this->wrapText = $wrapText;
        $this->style->mergeComponent('alignment', ['wrapText' => $wrapText]);

        return $this;
    }

    public function setShrinkToFit(bool $shrink): static
    {
        $this->shrinkToFit = $shrink;
        $this->style->mergeComponent('alignment', ['shrinkToFit' => $shrink]);

        return $this;
    }

    public function setTextRotation(int $rotation): static
    {
        $this->textRotation = $rotation;
        $this->style->mergeComponent('alignment', ['textRotation' => $rotation]);

        return $this;
    }

    public function setIndent(int $indent): static
    {
        $this->indent = $indent;
        $this->style->mergeComponent('alignment', ['indent' => $indent]);

        return $this;
    }

    public function getHorizontal(): string
    {
        return $this->horizontal;
    }

    public function getVertical(): string
    {
        return $this->vertical;
    }

    public function getWrapText(): bool
    {
        return $this->wrapText;
    }

    public function getShrinkToFit(): bool
    {
        return $this->shrinkToFit;
    }

    public function getTextRotation(): int
    {
        return $this->textRotation;
    }

    public function getIndent(): int
    {
        return $this->indent;
    }
}
