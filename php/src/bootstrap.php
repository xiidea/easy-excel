<?php

declare(strict_types=1);

/*
 * Registers lazy class aliases so existing code using
 * PhpOffice\PhpSpreadsheet\* transparently gets EasyExcel\Compat\*
 * (PLAN.md §14.1: alias bootstrap instead of composer-replace).
 *
 * The aliases are NOT installed when the real phpoffice/phpspreadsheet
 * package is present (it wins, easing incremental adoption) or when the
 * extension is unavailable, unless EASY_EXCEL_FORCE_ALIAS=1 forces them
 * (used by the test suite, which fakes the extension functions).
 */
(static function (): void {
    $force = (\getenv('EASY_EXCEL_FORCE_ALIAS') === '1');

    if (!$force) {
        if (!\function_exists('easy_excel_new')) {
            return; // extension missing: stay dormant
        }
        if (\class_exists(\Composer\InstalledVersions::class)
            && \Composer\InstalledVersions::isInstalled('phpoffice/phpspreadsheet')) {
            return; // real PhpSpreadsheet installed: defer to it
        }
    }

    \spl_autoload_register(static function (string $class): void {
        if (!\str_starts_with($class, 'PhpOffice\\PhpSpreadsheet\\')) {
            return;
        }
        $target = 'EasyExcel\\Compat\\' . \substr($class, 25);
        if (\class_exists($target) || \interface_exists($target)) {
            \class_alias($target, $class);
        }
    }, true, true);
})();
