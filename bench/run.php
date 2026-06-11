<?php

declare(strict_types=1);

/*
 * Benchmark runner. One library + workload per process so peak memory is
 * honest. Emits a CSV line:
 *
 *   lib,workload,rows,wall_seconds,peak_mem_bytes,file_bytes
 *
 * Usage:
 *   php run.php <lib> <workload> [rows]
 *     lib:      easy-excel | phpspreadsheet | openspout | fast-excel-writer | rap2hpoutre
 *     workload: write | read
 *
 * `write` produces rows x 10 mixed columns (string/int/float/date string).
 * `read` consumes the file produced by a prior `write` of the same lib
 * (or bench.xlsx if present), counting cells.
 */

require __DIR__ . '/vendor/autoload.php';

$lib = $argv[1] ?? null;
$workload = $argv[2] ?? 'write';
$rows = (int) ($argv[3] ?? 100_000);
$file = \sys_get_temp_dir() . "/bench-$lib.xlsx";

$adapters = require __DIR__ . '/adapters.php';
if (!isset($adapters[$lib])) {
    \fwrite(STDERR, 'usage: php run.php <' . \implode('|', \array_keys($adapters)) . "> <write|read> [rows]\n");
    exit(2);
}
if (!isset($adapters[$lib][$workload])) {
    \fwrite(STDERR, "$lib does not support workload $workload\n");
    exit(2);
}

/** @return iterable<list<mixed>> deterministic mixed-type rows, generated lazily */
function benchRows(int $count): iterable
{
    for ($i = 1; $i <= $count; ++$i) {
        yield [
            "customer-$i",
            'SKU-' . ($i % 1000),
            $i,
            \round($i * 1.37, 2),
            $i % 2 === 0 ? 'paid' : 'open',
            '2026-' . \str_pad((string) (($i % 12) + 1), 2, '0', STR_PAD_LEFT) . '-15',
            $i * 3,
            \round($i / 7, 4),
            'note text for row ' . $i,
            $i % 5,
        ];
    }
}

\gc_collect_cycles();
$t0 = \hrtime(true);
$result = $adapters[$lib][$workload]($file, $rows);
$wall = (\hrtime(true) - $t0) / 1e9;
$peak = \memory_get_peak_usage(true);
$size = \is_file($file) ? \filesize($file) : 0;

\printf("%s,%s,%d,%.3f,%d,%d\n", $lib, $workload, $workload === 'read' ? (int) $result : $rows, $wall, $peak, $size);
