<?php

declare(strict_types=1);

/*
 * Aliasing strategy for the PhpOffice\PhpSpreadsheet\* -> EasyExcel\Compat\*
 * bridge, factored into pure, unit-testable functions. bootstrap.php wires
 * these to the live environment; the test suite calls them directly.
 *
 * Modes (selected by aliasMode()):
 *
 *   off       Never alias. PhpOffice\PhpSpreadsheet\* resolves to the real
 *             package (or fatals if it isn't installed). Use this to run on
 *             stock PhpSpreadsheet, e.g. for A/B output comparison.
 *
 *   strict    All-or-nothing. Alias every Compat-implemented class; *throw*
 *             UnsupportedApiException on any PhpOffice\PhpSpreadsheet\* class
 *             the Compat layer lacks. A request is therefore served entirely
 *             by Compat or it fails — a handle-based workbook can never be
 *             mixed with a real object graph. This is the default when the
 *             native extension is loaded.
 *
 *   fallback  Hybrid escape hatch. Alias what Compat implements; defer every
 *             other PhpOffice\PhpSpreadsheet\* class to the real package
 *             (per class). Convenient for incremental adoption, but can mix
 *             object models within one request — opt in knowingly.
 */

namespace EasyExcel;

/**
 * Resolve the effective aliasing mode from environment + capability.
 *
 * @param string|false $env     Raw value of getenv('EASY_EXCEL_ALIAS')
 *                              (false when unset). Case-insensitive.
 * @param bool         $haveExt Whether the native easy_excel extension is loaded.
 *
 * @return 'off'|'strict'|'fallback'
 */
function aliasMode(string|false $env, bool $haveExt): string
{
    return match (\strtolower($env === false ? '' : $env)) {
        'off'      => 'off',
        'force'    => 'strict',            // force aliases even without the ext (tests)
        'strict'   => $haveExt ? 'strict' : 'off',
        'fallback' => $haveExt ? 'fallback' : 'off',
        default    => $haveExt ? 'strict' : 'off', // auto: prefer Compat when the ext is present
    };
}

/**
 * Map a class name to its Compat target.
 *
 * @return string|false|null
 *   string  FQN of the Compat class/interface to alias to
 *   false   a PhpOffice\PhpSpreadsheet\* class with no Compat implementation
 *   null    not a PhpOffice\PhpSpreadsheet\* class at all (ignore)
 */
function compatTarget(string $class): string|false|null
{
    if (!\str_starts_with($class, 'PhpOffice\\PhpSpreadsheet\\')) {
        return null;
    }
    $target = 'EasyExcel\\Compat\\' . \substr($class, 25);

    return (\class_exists($target) || \interface_exists($target)) ? $target : false;
}

/**
 * Decide what the autoloader should do for one class under a given mode.
 * Pure: no side effects, so every mode x class combination is unit-testable.
 *
 * @return array{0: 'ignore'|'alias'|'throw'|'defer', 1: ?string}
 *   ['ignore', null]      not a PhpOffice\PhpSpreadsheet\* class
 *   ['alias', <target>]   alias the class to the given Compat FQN
 *   ['throw', null]       strict mode, no Compat implementation -> fail loudly
 *   ['defer', null]       fallback mode, no Compat impl -> let the real package load it
 */
function aliasAction(string $mode, string $class): array
{
    $target = compatTarget($class);
    if ($target === null) {
        return ['ignore', null];
    }
    if ($target !== false) {
        return ['alias', $target];
    }

    return [$mode === 'strict' ? 'throw' : 'defer', null];
}

/**
 * Register the prepended autoloader that performs the aliasing for the given
 * mode. No-op for 'off'.
 */
function registerCompatAutoloader(string $mode): void
{
    if ($mode === 'off') {
        return;
    }

    \spl_autoload_register(static function (string $class) use ($mode): void {
        [$action, $target] = aliasAction($mode, $class);

        switch ($action) {
            case 'alias':
                \class_alias($target, $class);

                return;
            case 'throw':
                throw new UnsupportedApiException(\sprintf(
                    '%s is not implemented by the easy-excel Compat layer. Implement it, '
                    . 'or set EASY_EXCEL_ALIAS=off (or =fallback) to use real '
                    . 'phpoffice/phpspreadsheet for the whole request.',
                    $class,
                ));
            default: // 'ignore' | 'defer'
                return;
        }
    }, true, true);
}
