<?php

declare(strict_types=1);

namespace EasyExcel\Compat;

/**
 * Global settings, PhpSpreadsheet style (wave 5.2).
 *
 * PhpSpreadsheet uses this class to swap out pluggable back ends. easy-excel
 * has no such seams — rendering and calculation live in Go/excelize — so the
 * accessors here are state-only: they round-trip their value so consuming code
 * (and its tests) behave, but nothing reads them back out.
 *
 * `setChartRenderer()` is the one worth calling out: charts are emitted as
 * real Excel chart parts by excelize, never rasterized in PHP, so a renderer
 * such as JpGraph has nothing to do. Setting it is accepted and ignored
 * (COMPAT.md). It is deliberately *not* an exception — the call sites that use
 * it are guarding an HTML/PDF preview path, and throwing would break workbook
 * generation that is otherwise fully supported.
 */
final class Settings
{
    private static ?string $chartRenderer = null;

    private static string $locale = 'en';

    private static mixed $cache = null;

    private static mixed $httpClient = null;

    private static mixed $requestFactory = null;

    /**
     * Accepted and ignored: easy-excel writes native chart parts, so there is
     * no PHP-side rendering step for a renderer to service.
     */
    public static function setChartRenderer(string $rendererClass): void
    {
        self::$chartRenderer = $rendererClass;
    }

    public static function getChartRenderer(): ?string
    {
        return self::$chartRenderer;
    }

    public static function unsetChartRenderer(): void
    {
        self::$chartRenderer = null;
    }

    /**
     * Locale for formula translation and formatting.
     *
     * State-only: calculation is delegated to excelize, which has no locale
     * layer, so the value round-trips but does not change output. Returns true
     * for parity with upstream's success/failure contract.
     */
    public static function setLocale(string $locale): bool
    {
        self::$locale = $locale;

        return true;
    }

    public static function getLocale(): string
    {
        return self::$locale;
    }

    /** Matches upstream: the flag set used when escaping HTML output. */
    public static function htmlEntityFlags(): int
    {
        return \ENT_COMPAT;
    }

    /** Accepted no-op: the extension owns its own caching (COMPAT.md). */
    public static function setCache(mixed $cache): void
    {
        self::$cache = $cache;
    }

    public static function getCache(): mixed
    {
        return self::$cache;
    }

    public static function setHttpClient(mixed $httpClient, mixed $requestFactory = null): void
    {
        self::$httpClient = $httpClient;
        self::$requestFactory = $requestFactory;
    }

    public static function getHttpClient(): mixed
    {
        return self::$httpClient;
    }

    public static function getRequestFactory(): mixed
    {
        return self::$requestFactory;
    }

    public static function unsetHttpClient(): void
    {
        self::$httpClient = null;
        self::$requestFactory = null;
    }

    /** @internal test seam — resets every static back to its default */
    public static function reset(): void
    {
        self::$chartRenderer = null;
        self::$locale = 'en';
        self::$cache = null;
        self::$httpClient = null;
        self::$requestFactory = null;
    }
}
