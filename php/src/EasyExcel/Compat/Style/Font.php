<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/** Font component: setters forward partial specs through the parent Style. */
class Font
{
    public const UNDERLINE_NONE = 'none';
    public const UNDERLINE_SINGLE = 'single';
    public const UNDERLINE_DOUBLE = 'double';
    public const UNDERLINE_SINGLEACCOUNTING = 'singleAccounting';
    public const UNDERLINE_DOUBLEACCOUNTING = 'doubleAccounting';

    private bool $bold = false;
    private bool $italic = false;
    private bool $strikethrough = false;
    private string $underline = self::UNDERLINE_NONE;
    private string $name = 'Calibri';
    private float $size = 11.0;
    private ?Color $color = null;

    public function __construct(private Style $style)
    {
    }

    public function setBold(bool $bold): static
    {
        $this->bold = $bold;
        $this->style->mergeComponent('font', ['bold' => $bold]);

        return $this;
    }

    public function setItalic(bool $italic): static
    {
        $this->italic = $italic;
        $this->style->mergeComponent('font', ['italic' => $italic]);

        return $this;
    }

    public function setStrikethrough(bool $strike): static
    {
        $this->strikethrough = $strike;
        $this->style->mergeComponent('font', ['strikethrough' => $strike]);

        return $this;
    }

    public function setUnderline(bool|string $underline): static
    {
        if (\is_bool($underline)) {
            $underline = $underline ? self::UNDERLINE_SINGLE : self::UNDERLINE_NONE;
        }
        $this->underline = $underline;
        $this->style->mergeComponent('font', ['underline' => $underline]);

        return $this;
    }

    public function setName(string $name): static
    {
        $this->name = $name;
        $this->style->mergeComponent('font', ['name' => $name]);

        return $this;
    }

    public function setSize(mixed $size): static
    {
        $this->size = (float) $size;
        $this->style->mergeComponent('font', ['size' => (float) $size]);

        return $this;
    }

    public function setColor(Color $color): static
    {
        $this->color = $color;
        $this->style->mergeComponent('font', ['color' => ['argb' => $color->getARGB()]]);

        return $this;
    }

    public function getBold(): bool
    {
        return $this->bold;
    }

    public function getItalic(): bool
    {
        return $this->italic;
    }

    public function getStrikethrough(): bool
    {
        return $this->strikethrough;
    }

    public function getUnderline(): string
    {
        return $this->underline;
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getSize(): float
    {
        return $this->size;
    }

    public function getColor(): Color
    {
        return ($this->color ??= new Color())->bind(
            fn (string $argb) => $this->style->mergeComponent('font', ['color' => ['argb' => $argb]])
        );
    }
}
