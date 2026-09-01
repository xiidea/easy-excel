<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/**
 * Chart colour value, PhpSpreadsheet style (wave 5.4).
 *
 * A small holder: excelize takes plain colour strings, so this exists to give
 * Axis/GridLines the same setColorProperties() surface PhpSpreadsheet callers
 * use, and to normalise the value on the way through.
 */
class ChartColor
{
    public const EXCEL_COLOR_TYPE_ARGB = 'srgbClr';

    public const EXCEL_COLOR_TYPE_RGB = 'srgbClr';

    public const EXCEL_COLOR_TYPE_SCHEME = 'schemeClr';

    public const EXCEL_COLOR_TYPE_STANDARD = 'prstClr';

    private string $value = '';

    private string $type = '';

    private ?int $alpha = null;

    public function setColorProperties(?string $color, null|float|int|string $alpha = null, ?string $type = null): static
    {
        if ($color !== null) {
            $this->value = self::normalise($color);
        }
        if ($type !== null) {
            $this->type = $type;
        }
        $this->alpha = $alpha === null ? null : (int) $alpha;

        return $this;
    }

    public function getValue(): string
    {
        return $this->value;
    }

    public function getType(): string
    {
        return $this->type;
    }

    public function getAlpha(): ?int
    {
        return $this->alpha;
    }

    /**
     * excelize expects a bare RGB/ARGB hex string. PhpSpreadsheet callers pass
     * either form, with or without a leading '#'.
     */
    public static function normalise(string $color): string
    {
        $color = \ltrim($color, '#');

        return \strtoupper($color);
    }
}
