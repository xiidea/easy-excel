<?php

declare(strict_types=1);

/*
 * Phase 1: PhpSpreadsheet -> Compat surface diff.
 *
 * Enumerates the public API surface of a real phpoffice/phpspreadsheet install
 * (every class/interface, plus their public methods and constants) and diffs it
 * against the EasyExcel\Compat layer. This quantifies the all-or-nothing gap and
 * is meant to run in CI as a gate: bump the baseline deliberately, and any *new*
 * uncovered class/method/constant fails the build instead of surfacing in
 * production (cf. the PAPERSIZE_ and PAGEORDER_ constants found at runtime).
 *
 * Usage:
 *   php tools/compat-surface-diff.php [PHPSPREADSHEET_SRC] [options]
 *
 *   PHPSPREADSHEET_SRC   Path to .../phpoffice/phpspreadsheet/src/PhpSpreadsheet
 *                        (auto-discovered under php/vendor if omitted).
 *
 * Options:
 *   --members            Also diff public methods and constants of covered classes.
 *   --json               Emit the full report as JSON instead of text.
 *   --baseline=FILE      Compare missing-class set against an allow-list JSON;
 *                        exit 1 only on entries not present in it (regressions).
 *   --update-baseline=FILE  Write the current missing-class set to FILE and exit 0.
 *
 * Exit codes: 0 = no new gaps (or report-only), 1 = new gaps vs baseline,
 *             2 = could not locate a PhpSpreadsheet install.
 */

const PS_PREFIX = 'PhpOffice\\PhpSpreadsheet\\';
const COMPAT_PREFIX = 'EasyExcel\\Compat\\';

// We reflect real PhpOffice\PhpSpreadsheet\* and EasyExcel\Compat\* side by
// side; aliasing one onto the other would defeat the diff.
\putenv('EASY_EXCEL_ALIAS=off');

$argvRest = \array_slice($argv, 1);
$opts = ['members' => false, 'json' => false, 'baseline' => null, 'update' => null, 'autoload' => null];
$srcArg = null;
foreach ($argvRest as $a) {
    if ($a === '--members') {
        $opts['members'] = true;
    } elseif ($a === '--json') {
        $opts['json'] = true;
    } elseif (\str_starts_with($a, '--baseline=')) {
        $opts['baseline'] = \substr($a, 11);
    } elseif (\str_starts_with($a, '--update-baseline=')) {
        $opts['update'] = \substr($a, 18);
    } elseif (\str_starts_with($a, '--autoload=')) {
        $opts['autoload'] = \substr($a, 11);
    } elseif (!\str_starts_with($a, '--')) {
        $srcArg = $a;
    } else {
        \fwrite(\STDERR, "unknown option: $a\n");
        exit(2);
    }
}

$psSrc = locatePhpSpreadsheet($srcArg);
if ($psSrc === null) {
    \fwrite(\STDERR,
        "Could not find a phpoffice/phpspreadsheet install.\n" .
        "Pass its src path explicitly, e.g.:\n" .
        "  composer require --dev phpoffice/phpspreadsheet\n" .
        "  php tools/compat-surface-diff.php vendor/phpoffice/phpspreadsheet/src/PhpSpreadsheet\n");
    exit(2);
}

// Prefer Composer's autoloader when present: it resolves PhpSpreadsheet, its
// dependencies, *and* the Compat layer, so reflection sees the full surface.
// Otherwise fall back to path-based PSR-4 for the no-Composer (local) case.
$autoload = $opts['autoload'] ?? defaultComposerAutoload();
if ($autoload !== null && \is_file($autoload)) {
    require $autoload;
} else {
    registerPsr4(COMPAT_PREFIX, \dirname(__DIR__) . '/src/EasyExcel/Compat');
}
registerPsr4(PS_PREFIX, $psSrc); // path fallback if not already covered by Composer

$report = buildReport($psSrc, $opts['members']);

if ($opts['update'] !== null) {
    $json = \json_encode($report['missingClasses'], \JSON_PRETTY_PRINT | \JSON_UNESCAPED_SLASHES) . "\n";
    $dest = $opts['update'] === '-' ? 'php://stdout' : $opts['update'];
    \file_put_contents($dest, $json);
    \fwrite(\STDERR, \sprintf("baseline: %d missing classes -> %s\n",
        \count($report['missingClasses']), $opts['update'] === '-' ? 'stdout' : $opts['update']));
    exit(0);
}

$regressions = [];
if ($opts['baseline'] !== null) {
    $allowed = \is_file($opts['baseline'])
        ? (array) \json_decode((string) \file_get_contents($opts['baseline']), true)
        : [];
    $regressions = \array_values(\array_diff($report['missingClasses'], $allowed));
}

if ($opts['json']) {
    echo \json_encode(['report' => $report, 'regressions' => $regressions],
        \JSON_PRETTY_PRINT | \JSON_UNESCAPED_SLASHES), "\n";
} else {
    printText($report, $opts, $regressions);
}

exit($regressions === [] ? 0 : 1);

// ---------------------------------------------------------------------------

function locatePhpSpreadsheet(?string $arg): ?string
{
    $candidates = \array_filter([
        $arg,
        \dirname(__DIR__) . '/vendor/phpoffice/phpspreadsheet/src/PhpSpreadsheet',
        \dirname(__DIR__, 2) . '/vendor/phpoffice/phpspreadsheet/src/PhpSpreadsheet',
    ]);
    foreach ($candidates as $c) {
        if (\is_dir($c) && \is_file($c . '/Spreadsheet.php')) {
            return \rtrim($c, '/');
        }
    }

    return null;
}

function defaultComposerAutoload(): ?string
{
    $path = \dirname(__DIR__) . '/vendor/autoload.php';

    return \is_file($path) ? $path : null;
}

function registerPsr4(string $prefix, string $baseDir): void
{
    \spl_autoload_register(static function (string $class) use ($prefix, $baseDir): void {
        if (!\str_starts_with($class, $prefix)) {
            return;
        }
        $rel = \str_replace('\\', '/', \substr($class, \strlen($prefix)));
        $file = $baseDir . '/' . $rel . '.php';
        if (\is_file($file)) {
            require $file;
        }
    });
}

/**
 * @return array{
 *   total:int, covered:int, missing:int, coverage:float,
 *   missingClasses:list<string>,
 *   memberGaps:array<string, array{methods:list<string>, constants:list<string>}>
 * }
 */
function buildReport(string $psSrc, bool $members): array
{
    $missingClasses = [];
    $memberGaps = [];
    $total = 0;
    $covered = 0;

    foreach (psTypes($psSrc) as $fqn) {
        ++$total;
        $target = COMPAT_PREFIX . \substr($fqn, \strlen(PS_PREFIX));

        if (!\class_exists($target) && !\interface_exists($target)) {
            $missingClasses[] = $fqn;

            continue;
        }
        ++$covered;

        if ($members) {
            $gap = memberGap($fqn, $target);
            if ($gap['methods'] !== [] || $gap['constants'] !== []) {
                $memberGaps[$fqn] = $gap;
            }
        }
    }

    \sort($missingClasses);
    \ksort($memberGaps);

    return [
        'total' => $total,
        'covered' => $covered,
        'missing' => \count($missingClasses),
        'coverage' => $total > 0 ? \round($covered / $total * 100, 1) : 0.0,
        'missingClasses' => $missingClasses,
        'memberGaps' => $memberGaps,
    ];
}

/** @return iterable<string> fully-qualified PhpSpreadsheet class/interface names */
function psTypes(string $psSrc): iterable
{
    $it = new RecursiveIteratorIterator(
        new RecursiveDirectoryIterator($psSrc, FilesystemIterator::SKIP_DOTS),
    );
    foreach ($it as $file) {
        if ($file->getExtension() !== 'php') {
            continue;
        }
        $rel = \substr($file->getPathname(), \strlen($psSrc) + 1, -4); // strip base + ".php"
        $fqn = PS_PREFIX . \str_replace('/', '\\', $rel);

        try {
            if (!\class_exists($fqn) && !\interface_exists($fqn)) {
                continue; // not a type declaration (or failed to load) — skip
            }
            $rc = new ReflectionClass($fqn);
            if ($rc->isInternal() || $rc->isTrait()) {
                continue;
            }
        } catch (\Throwable) {
            continue; // unreflectable (missing dep, etc.) — out of scope
        }

        yield $fqn;
    }
}

/** @return array{methods:list<string>, constants:list<string>} members PS declares that Compat lacks */
function memberGap(string $psFqn, string $compatFqn): array
{
    $ps = new ReflectionClass($psFqn);
    $compat = new ReflectionClass($compatFqn);

    $missingMethods = [];
    foreach ($ps->getMethods(ReflectionMethod::IS_PUBLIC) as $m) {
        if ($m->getDeclaringClass()->getName() !== $psFqn) {
            continue; // only members PS itself declares
        }
        if ($m->isStatic() && $m->getName() === '__set_state') {
            continue;
        }
        if (!$compat->hasMethod($m->getName())) {
            $missingMethods[] = $m->getName();
        }
    }

    $compatConsts = $compat->getConstants();
    $missingConsts = [];
    foreach ($ps->getReflectionConstants(ReflectionClassConstant::IS_PUBLIC) as $c) {
        if ($c->getDeclaringClass()->getName() !== $psFqn) {
            continue;
        }
        if (!\array_key_exists($c->getName(), $compatConsts)) {
            $missingConsts[] = $c->getName();
        }
    }

    \sort($missingMethods);
    \sort($missingConsts);

    return ['methods' => $missingMethods, 'constants' => $missingConsts];
}

function printText(array $report, array $opts, array $regressions): void
{
    echo "PhpSpreadsheet -> Compat surface diff\n";
    echo \str_repeat('=', 40), "\n";
    echo \sprintf("classes/interfaces : %d\n", $report['total']);
    echo \sprintf("covered by Compat  : %d (%.1f%%)\n", $report['covered'], $report['coverage']);
    echo \sprintf("missing            : %d\n\n", $report['missing']);

    if ($report['missingClasses'] !== []) {
        echo "Uncovered classes (strict mode throws on these):\n";
        foreach ($report['missingClasses'] as $c) {
            echo "  - $c\n";
        }
        echo "\n";
    }

    if ($opts['members'] && $report['memberGaps'] !== []) {
        echo "Member gaps in covered classes:\n";
        foreach ($report['memberGaps'] as $class => $gap) {
            echo "  $class\n";
            foreach ($gap['constants'] as $const) {
                echo "      const  $const\n";
            }
            foreach ($gap['methods'] as $method) {
                echo "      method $method()\n";
            }
        }
        echo "\n";
    }

    if ($opts['baseline'] !== null) {
        if ($regressions === []) {
            echo "No new uncovered classes vs baseline. OK\n";
        } else {
            echo \sprintf("%d NEW uncovered class(es) vs baseline:\n", \count($regressions));
            foreach ($regressions as $c) {
                echo "  + $c\n";
            }
        }
    }
}
