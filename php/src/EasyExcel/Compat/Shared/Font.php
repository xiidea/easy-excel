<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Shared;

/**
 * Font metrics, PhpSpreadsheet style (wave 5.2).
 *
 * PhpSpreadsheet measures rendered text with GD/afm metrics to size columns.
 * easy-excel does not: auto-size is approximated in Go at save time
 * (COMPAT.md divergence 10). What is implemented here is the part that is pure
 * table lookup and is genuinely needed by callers laying out HTML — default
 * row height and character width per font size.
 *
 * The $font parameters are typed `?object` rather than Style\Font on purpose:
 * Compat's Style\Font is bound to its owning Style and cannot be constructed
 * standalone, while PhpSpreadsheet callers pass whatever font-ish object they
 * hold. Anything exposing getName()/getSize() works; anything else falls back
 * to the Calibri 11 defaults instead of fataling.
 */
final class Font
{
    public const AUTOSIZE_METHOD_APPROX = 'approx';

    public const AUTOSIZE_METHOD_EXACT = 'exact';

    /** Default character width in pixels for 11pt Calibri, per OOXML. */
    private const DEFAULT_CHARACTER_WIDTH = 7;

    /**
     * Default row height in points for a given font, from PhpSpreadsheet's
     * table for the common families; other families fall back to a linear
     * approximation of the same curve.
     */
    public static function getDefaultRowHeightByFont(?object $font): float
    {
        $name = \is_callable([$font, 'getName']) ? (string) $font->getName() : 'Calibri';
        $size = \is_callable([$font, 'getSize']) ? (float) $font->getSize() : 11.0;

        $table = [
            'Arial' => [10 => 12.75, 9 => 12.0, 8 => 11.25, 11 => 14.25, 12 => 15.0],
            'Calibri' => [11 => 15.0, 10 => 12.75, 9 => 12.0, 8 => 11.25, 12 => 15.75],
            'Verdana' => [10 => 12.75, 9 => 11.25, 8 => 10.5, 11 => 14.25],
            'Times New Roman' => [10 => 12.75, 9 => 12.0, 8 => 11.25, 11 => 14.25],
        ];

        if (isset($table[$name][(int) $size])) {
            return $table[$name][(int) $size];
        }

        // Same slope PhpSpreadsheet's table follows, for unlisted combinations.
        return \round($size * 1.3636, 2);
    }

    /**
     * Approximate character width in pixels for the font, used to convert
     * Excel column-width units to pixels.
     */
    public static function getCharacterWidth(?object $font): int
    {
        if (!\is_callable([$font, 'getSize'])) {
            return self::DEFAULT_CHARACTER_WIDTH;
        }
        $size = (float) ($font->getSize() ?: 11);
        // 7px at 11pt, scaled linearly — matches the OOXML default grid.
        $width = (int) \round(self::DEFAULT_CHARACTER_WIDTH * $size / 11);

        return \max(1, $width);
    }

    /**
     * Not supported: exact text measurement needs the font metrics easy-excel
     * deliberately does not carry. Returns null so callers fall back to their
     * own estimate, which is what PhpSpreadsheet does when metrics are absent.
     */
    public static function getAutoSizeMethod(): string
    {
        return self::AUTOSIZE_METHOD_APPROX;
    }

    /** Accepted no-op: auto-size is approximated in Go (COMPAT.md §10). */
    public static function setAutoSizeMethod(string $method): bool
    {
        return $method === self::AUTOSIZE_METHOD_APPROX;
    }
}
