<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Shared;

/**
 * String helpers, PhpSpreadsheet style (wave 5.1).
 *
 * Pure PHP — no extension involvement. Scoped to the members real consumers
 * call: the writers use these for column-letter arithmetic, margin formatting
 * and control-character escaping.
 */
final class StringHelper
{
    /** Locale decimal separator, as PhpSpreadsheet exposes it. */
    private static ?string $decimalSeparator = null;

    /** Locale thousands separator. */
    private static ?string $thousandsSeparator = null;

    /**
     * Coerce a value to string the way PhpSpreadsheet's writers expect.
     *
     * @param bool $throw        throw on a non-stringable value instead of
     *                           returning $default
     * @param bool $convertBool  render booleans as TRUE/FALSE rather than 1/''
     */
    public static function convertToString(
        mixed $value,
        bool $throw = true,
        string $default = '',
        bool $convertBool = false,
        bool $lessFloatPrecision = false,
    ): string {
        if ($convertBool && \is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }
        if ($value === null || \is_scalar($value)) {
            if ($lessFloatPrecision && \is_float($value)) {
                return (string) \round($value, 10);
            }

            return (string) $value;
        }
        if ($value instanceof \Stringable) {
            return (string) $value;
        }
        if ($throw) {
            throw new \TypeError('easy-excel: value cannot be converted to string');
        }

        return $default;
    }

    /**
     * Increment a spreadsheet column string in place: A → B, Z → AA, AZ → BA.
     *
     * Uses str_increment() where available: bare ++ on a non-numeric string is
     * deprecated as of PHP 8.3 and would emit a notice on every column step.
     *
     * @param string $str modified in place, and also returned
     */
    public static function stringIncrement(string &$str): string
    {
        $str = $str === '' ? 'A' : (\function_exists('str_increment') ? \str_increment($str) : ++$str);

        return $str;
    }

    /**
     * Format a number for CSS/HTML output — always a plain decimal with '.' as
     * the separator, never locale-formatted or in exponent notation.
     */
    public static function formatNumber(float|int|string|null $numericValue): string
    {
        // Upstream returns '' (not '0') for null/empty — matched deliberately:
        // the writers concatenate this straight into CSS margin declarations.
        if ($numericValue === null || $numericValue === '') {
            return '';
        }
        if (\is_string($numericValue) && !\is_numeric($numericValue)) {
            return $numericValue;
        }

        return (string) (float) $numericValue === (string) (int) $numericValue
            ? (string) (int) $numericValue
            : \rtrim(\rtrim(\number_format((float) $numericValue, 10, '.', ''), '0'), '.');
    }

    /**
     * Escape control characters into the _xHHHH_ form OOXML requires.
     */
    public static function controlCharacterPHP2OOXML(string $textValue): string
    {
        return (string) \preg_replace_callback(
            '/[\x00-\x08\x0b\x0c\x0e-\x1f]/',
            static fn (array $m): string => \sprintf('_x%04X_', \ord($m[0])),
            $textValue,
        );
    }

    /** Reverse of controlCharacterPHP2OOXML. */
    public static function controlCharacterOOXML2PHP(string $textValue): string
    {
        return (string) \preg_replace_callback(
            '/_x([0-9A-Fa-f]{4})_/',
            static fn (array $m): string => \chr((int) \hexdec($m[1])),
            $textValue,
        );
    }

    /** True when the string is valid UTF-8. */
    public static function isUTF8(string $textValue): bool
    {
        return $textValue === '' || \preg_match('/^./su', $textValue) === 1;
    }

    /** Substring that is safe on multibyte input. */
    public static function substring(string $textValue, int $offset, ?int $length = null): string
    {
        return \mb_substr($textValue, $offset, $length, 'UTF-8');
    }

    public static function countCharacters(string $textValue, string $encoding = 'UTF-8'): int
    {
        return \mb_strlen($textValue, $encoding);
    }

    public static function strToUpper(string $textValue): string
    {
        return \mb_convert_case($textValue, \MB_CASE_UPPER, 'UTF-8');
    }

    public static function strToLower(string $textValue): string
    {
        return \mb_convert_case($textValue, \MB_CASE_LOWER, 'UTF-8');
    }

    public static function strToTitle(string $textValue): string
    {
        return \mb_convert_case($textValue, \MB_CASE_TITLE, 'UTF-8');
    }

    public static function getDecimalSeparator(): string
    {
        if (self::$decimalSeparator === null) {
            $localeconv = \localeconv();
            self::$decimalSeparator = ($localeconv['decimal_point'] !== '' )
                ? $localeconv['decimal_point']
                : '.';
        }

        return self::$decimalSeparator;
    }

    public static function setDecimalSeparator(string $separator): void
    {
        self::$decimalSeparator = $separator;
    }

    public static function getThousandsSeparator(): string
    {
        if (self::$thousandsSeparator === null) {
            $localeconv = \localeconv();
            self::$thousandsSeparator = ($localeconv['thousands_sep'] !== '')
                ? $localeconv['thousands_sep']
                : ',';
        }

        return self::$thousandsSeparator;
    }

    public static function setThousandsSeparator(string $separator): void
    {
        self::$thousandsSeparator = $separator;
    }

    /** @internal test seam */
    public static function reset(): void
    {
        self::$decimalSeparator = null;
        self::$thousandsSeparator = null;
    }
}
