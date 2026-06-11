<?php

declare(strict_types=1);

/*
 * Zero-dependency test runner: discovers tests/cases/*.php, each returning
 * array<string, callable>. Runs everywhere PHP runs — no composer install
 * needed. The PHPUnit-based PhpSpreadsheet compatibility suite is a separate
 * Phase-1-exit deliverable (PLAN.md §12).
 */

\putenv('EASY_EXCEL_FORCE_ALIAS=1');

// PSR-4 for EasyExcel\ without composer
\spl_autoload_register(static function (string $class): void {
    if (\str_starts_with($class, 'EasyExcel\\')) {
        $path = __DIR__ . '/../src/EasyExcel/' . \str_replace('\\', '/', \substr($class, 10)) . '.php';
        if (\is_file($path)) {
            require $path;
        }
    }
});

require __DIR__ . '/Fake/functions.php';
require __DIR__ . '/../src/bootstrap.php';

final class T
{
    public static int $assertions = 0;

    public static function same(mixed $expected, mixed $actual, string $msg = ''): void
    {
        ++self::$assertions;
        if ($expected !== $actual) {
            throw new RuntimeException(\sprintf(
                "%s\n  expected: %s\n  actual:   %s",
                $msg !== '' ? $msg : 'values differ',
                \var_export($expected, true),
                \var_export($actual, true),
            ));
        }
    }

    public static function ok(bool $cond, string $msg): void
    {
        ++self::$assertions;
        if (!$cond) {
            throw new RuntimeException($msg);
        }
    }

    public static function throws(string $class, callable $fn, string $msg = ''): void
    {
        ++self::$assertions;
        try {
            $fn();
        } catch (\Throwable $e) {
            if ($e instanceof $class) {
                return;
            }
            throw new RuntimeException(($msg ?: 'wrong exception type') . ': got ' . $e::class . ': ' . $e->getMessage());
        }
        throw new RuntimeException($msg ?: "expected $class, nothing thrown");
    }
}

$failures = 0;
$passed = 0;
foreach (\glob(__DIR__ . '/cases/*.php') as $file) {
    $tests = require $file;
    foreach ($tests as $name => $fn) {
        EasyExcelFake::reset();
        try {
            $fn();
            ++$passed;
            echo "  ok  $name\n";
        } catch (\Throwable $e) {
            ++$failures;
            echo "FAIL  $name\n      {$e->getMessage()}\n";
            foreach (\array_slice(\explode("\n", $e->getTraceAsString()), 0, 3) as $line) {
                echo "      $line\n";
            }
        }
    }
}

echo "\n$passed passed, $failures failed, " . T::$assertions . " assertions\n";
exit($failures > 0 ? 1 : 0);
