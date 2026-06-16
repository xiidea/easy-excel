<?php

declare(strict_types=1);

/*
 * Installs the lazy class aliases that let existing code using
 * PhpOffice\PhpSpreadsheet\* transparently get EasyExcel\Compat\*
 * (PLAN.md §14.1: alias bootstrap instead of composer-replace).
 *
 * Precedence is driven by aliasMode() (see aliasing.php):
 *
 *   - extension loaded  -> "strict" all-or-nothing: Compat owns the whole
 *                          PhpOffice\PhpSpreadsheet\* namespace; an unimplemented
 *                          class throws UnsupportedApiException instead of
 *                          silently mixing with a real object graph.
 *   - extension missing -> "off": defer entirely to real phpoffice/phpspreadsheet.
 *
 * Override with EASY_EXCEL_ALIAS = off | strict | fallback | force. The legacy
 * EASY_EXCEL_FORCE_ALIAS=1 flag still maps to "force" (used by the test suite,
 * which fakes the extension functions).
 */

require_once __DIR__ . '/aliasing.php';

(static function (): void {
    $env = \getenv('EASY_EXCEL_ALIAS');
    if ($env === false && \getenv('EASY_EXCEL_FORCE_ALIAS') === '1') {
        $env = 'force'; // backward-compatible alias for the old flag
    }

    $mode = \EasyExcel\aliasMode($env, \function_exists('easy_excel_new'));
    \EasyExcel\registerCompatAutoloader($mode);
})();
