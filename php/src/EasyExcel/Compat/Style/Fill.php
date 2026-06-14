<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Style;

/** Fill component (pattern fills; gradients are unsupported, see COMPAT.md). */
class Fill
{
    public const FILL_NONE = 'none';
    public const FILL_SOLID = 'solid';
    public const FILL_PATTERN_MEDIUMGRAY = 'mediumGray';
    public const FILL_PATTERN_DARKGRAY = 'darkGray';
    public const FILL_PATTERN_LIGHTGRAY = 'lightGray';
    public const FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal';
    public const FILL_PATTERN_DARKVERTICAL = 'darkVertical';
    public const FILL_PATTERN_DARKDOWN = 'darkDown';
    public const FILL_PATTERN_DARKUP = 'darkUp';
    public const FILL_PATTERN_DARKGRID = 'darkGrid';
    public const FILL_PATTERN_DARKTRELLIS = 'darkTrellis';
    public const FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal';
    public const FILL_PATTERN_LIGHTVERTICAL = 'lightVertical';
    public const FILL_PATTERN_LIGHTDOWN = 'lightDown';
    public const FILL_PATTERN_LIGHTUP = 'lightUp';
    public const FILL_PATTERN_LIGHTGRID = 'lightGrid';
    public const FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis';
    public const FILL_PATTERN_GRAY125 = 'gray125';
    public const FILL_PATTERN_GRAY0625 = 'gray0625';
    public const FILL_GRADIENT_LINEAR = 'linear';
    public const FILL_GRADIENT_PATH = 'path';

    private ?string $fillType = null;
    private ?Color $startColor = null;
    private ?Color $endColor = null;

    public function __construct(private Style $style)
    {
    }

    public function setFillType(string $fillType): static
    {
        $this->fillType = $fillType;
        $this->style->mergeComponent('fill', ['fillType' => $fillType]);

        return $this;
    }

    private float $rotation = 0.0;

    /** Gradient angle in degrees (bucketed to a shading direction, COMPAT.md). */
    public function setRotation(float $rotation): static
    {
        $this->rotation = $rotation;
        $this->style->mergeComponent('fill', ['rotation' => $rotation]);

        return $this;
    }

    public function getRotation(): float
    {
        return $this->rotation;
    }

    public function getFillType(): string
    {
        return $this->fillType
            ?? (string) ($this->style->nativeComponent('fill')['fillType'] ?? self::FILL_NONE);
    }

    public function setStartColor(Color $color): static
    {
        $this->startColor = $color;
        $this->style->mergeComponent('fill', ['startColor' => ['argb' => $color->getARGB()]]);

        return $this;
    }

    public function setEndColor(Color $color): static
    {
        $this->endColor = $color;
        $this->style->mergeComponent('fill', ['endColor' => ['argb' => $color->getARGB()]]);

        return $this;
    }

    public function getStartColor(): Color
    {
        return ($this->startColor ??= $this->nativeColor('startColor'))->bind(
            fn (string $argb) => $this->style->mergeComponent('fill', ['startColor' => ['argb' => $argb]])
        );
    }

    public function getEndColor(): Color
    {
        return ($this->endColor ??= $this->nativeColor('endColor'))->bind(
            fn (string $argb) => $this->style->mergeComponent('fill', ['endColor' => ['argb' => $argb]])
        );
    }

    /** local writes win; otherwise read the color back from the stylesheet */
    private function nativeColor(string $key): Color
    {
        $native = $this->style->nativeComponent('fill')[$key] ?? null;
        if (\is_array($native)) {
            $argb = $native['argb'] ?? (isset($native['rgb']) ? 'FF' . $native['rgb'] : null);
            if ($argb !== null) {
                return new Color($argb);
            }
        }

        return new Color();
    }
}
