<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/** One border side; mutations are forwarded keyed by the side's position. */
class Border
{
    public const BORDER_NONE = 'none';
    public const BORDER_THIN = 'thin';
    public const BORDER_MEDIUM = 'medium';
    public const BORDER_DASHED = 'dashed';
    public const BORDER_DOTTED = 'dotted';
    public const BORDER_THICK = 'thick';
    public const BORDER_DOUBLE = 'double';
    public const BORDER_HAIR = 'hair';
    public const BORDER_MEDIUMDASHED = 'mediumDashed';
    public const BORDER_DASHDOT = 'dashDot';
    public const BORDER_MEDIUMDASHDOT = 'mediumDashDot';
    public const BORDER_DASHDOTDOT = 'dashDotDot';
    public const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    public const BORDER_SLANTDASHDOT = 'slantDashDot';

    private string $borderStyle = self::BORDER_NONE;
    private ?Color $color = null;

    public function __construct(private Style $style, private string $position)
    {
    }

    public function setBorderStyle(bool|string $borderStyle): static
    {
        if (\is_bool($borderStyle)) {
            $borderStyle = $borderStyle ? self::BORDER_MEDIUM : self::BORDER_NONE;
        }
        $this->borderStyle = $borderStyle;
        $this->style->mergeComponent('borders', [$this->position => ['borderStyle' => $borderStyle]]);

        return $this;
    }

    public function getBorderStyle(): string
    {
        return $this->borderStyle;
    }

    public function setColor(Color $color): static
    {
        $this->color = $color;
        $this->style->mergeComponent('borders', [$this->position => ['color' => ['argb' => $color->getARGB()]]]);

        return $this;
    }

    public function getColor(): Color
    {
        return ($this->color ??= new Color())->bind(
            fn (string $argb) => $this->style->mergeComponent('borders', [$this->position => ['color' => ['argb' => $argb]]])
        );
    }
}
