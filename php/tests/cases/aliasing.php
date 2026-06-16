<?php

declare(strict_types=1);

use EasyExcel\UnsupportedApiException;

use function EasyExcel\aliasAction;
use function EasyExcel\aliasMode;
use function EasyExcel\compatTarget;

/*
 * Phase 0: the all-or-nothing aliasing switch. The pure decision functions
 * are exercised across every mode; the live (strict) autoloader installed by
 * the test bootstrap is exercised by referencing an unimplemented class.
 */

return [
    'aliasMode: auto prefers Compat when the extension is loaded' => function (): void {
        T::same('strict', aliasMode(false, true), 'auto + ext -> strict');
        T::same('off', aliasMode(false, false), 'auto, no ext -> off');
    },

    'aliasMode: off always defers to real PhpSpreadsheet' => function (): void {
        T::same('off', aliasMode('off', true));
        T::same('off', aliasMode('OFF', false), 'case-insensitive');
    },

    'aliasMode: force enables strict even without the extension' => function (): void {
        T::same('strict', aliasMode('force', false));
    },

    'aliasMode: strict/fallback need the extension, else degrade to off' => function (): void {
        T::same('strict', aliasMode('strict', true));
        T::same('off', aliasMode('strict', false), 'strict without ext -> off');
        T::same('fallback', aliasMode('fallback', true));
        T::same('off', aliasMode('fallback', false), 'fallback without ext -> off');
    },

    'compatTarget: maps implemented classes, flags gaps, ignores foreign ns' => function (): void {
        T::same('EasyExcel\\Compat\\Spreadsheet', compatTarget('PhpOffice\\PhpSpreadsheet\\Spreadsheet'));
        T::same(false, compatTarget('PhpOffice\\PhpSpreadsheet\\Writer\\Ods'), 'unimplemented -> false');
        T::same(null, compatTarget('Some\\Other\\Library'), 'foreign namespace -> null');
    },

    'aliasAction: strict aliases implemented, throws on gaps' => function (): void {
        T::same(['alias', 'EasyExcel\\Compat\\Spreadsheet'],
            aliasAction('strict', 'PhpOffice\\PhpSpreadsheet\\Spreadsheet'));
        T::same(['throw', null],
            aliasAction('strict', 'PhpOffice\\PhpSpreadsheet\\Writer\\Ods'),
            'strict must fail loudly on an unimplemented class');
        T::same(['ignore', null],
            aliasAction('strict', 'Some\\Other\\Library'));
    },

    'aliasAction: fallback defers gaps instead of throwing' => function (): void {
        T::same(['alias', 'EasyExcel\\Compat\\Spreadsheet'],
            aliasAction('fallback', 'PhpOffice\\PhpSpreadsheet\\Spreadsheet'));
        T::same(['defer', null],
            aliasAction('fallback', 'PhpOffice\\PhpSpreadsheet\\Writer\\Ods'),
            'fallback hands unimplemented classes to the real package');
    },

    'strict autoloader throws UnsupportedApiException for an uncovered class' => function (): void {
        // the test bootstrap runs in strict mode (EASY_EXCEL_FORCE_ALIAS=1)
        T::throws(UnsupportedApiException::class, static function (): void {
            \class_exists('PhpOffice\\PhpSpreadsheet\\Writer\\Ods');
        }, 'referencing an unimplemented PhpOffice class must fail loudly in strict mode');
    },

    'strict autoloader still aliases an implemented class' => function (): void {
        T::ok(
            \class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet'),
            'implemented classes keep resolving through the alias',
        );
    },
];
