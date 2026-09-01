<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Shared;

/**
 * Unit conversions for drawings and dimensions, PhpSpreadsheet style
 * (wave 5.2). Pure arithmetic — no extension involvement — so the ratios below
 * are the same constants PhpSpreadsheet uses (96 DPI screen, 72 pt/inch).
 */
final class Drawing
{
    /** Points → pixels at 96 DPI. */
    public static function pointsToPixels(int|float $points): int
    {
        return $points == 0 ? 0 : (int) \ceil($points * 96 / 72);
    }

    /** Pixels → points at 96 DPI. */
    public static function pixelsToPoints(int|float $pixels): float
    {
        return $pixels * 72 / 96;
    }

    /** Pixels → EMU (English Metric Units), 9525 EMU per pixel. */
    public static function pixelsToEMU(int|float $pixels): int
    {
        return (int) \round($pixels * 9525);
    }

    /** EMU → pixels. */
    public static function EMUToPixels(int|float $emu): int
    {
        return $emu == 0 ? 0 : (int) \round($emu / 9525);
    }

    /** Centimetres → EMU. */
    public static function centimetersToEMU(int|float $cm): int
    {
        return (int) \round($cm * 360000);
    }

    /** Inches → EMU. */
    public static function inchesToEMU(int|float $inches): int
    {
        return (int) \round($inches * 914400);
    }

    /** Degrees → the 60000ths-of-a-degree unit OOXML uses for rotation. */
    public static function degreesToAngle(int|float $degrees): int
    {
        return (int) \round($degrees * 60000);
    }

    public static function angleToDegrees(int|float $angle): int
    {
        return $angle == 0 ? 0 : (int) \round($angle / 60000);
    }

    /**
     * Column width (in Excel's character units) → pixels, using the font's
     * character width. Mirrors PhpSpreadsheet: a negative width means "not
     * set" and yields the font's default column width in pixels.
     */
    public static function cellDimensionToPixels(int|float $width, ?object $font = null): int
    {
        if ($width < 0) {
            return 0;
        }
        $charWidth = Font::getCharacterWidth($font);

        return (int) \round($width * $charWidth);
    }

    /** Pixels → column width in Excel's character units. */
    public static function pixelsToCellDimension(int|float $pixels, ?object $font = null): float
    {
        $charWidth = Font::getCharacterWidth($font);

        return $charWidth == 0 ? 0.0 : $pixels / $charWidth;
    }
}
