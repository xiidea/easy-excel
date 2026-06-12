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
 * `write-styled` is `write` plus a styled header, two column number formats,
 * column widths, freeze pane and auto-filter — the typical report
 * (phpspreadsheet and easy-excel only; both run the identical code through
 * the PhpSpreadsheet API surface).
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

/**
 * Identical styled-report workload across PhpSpreadsheet-compatible APIs;
 * $spreadsheet/$writer come from either library.
 */
function benchStyledWrite(object $spreadsheet, string $writerClass, string $file, int $rows): void
{
    $ws = $spreadsheet->getActiveSheet();
    $last = $rows + 1;
    $ws->getStyle('A1:J1')->applyFromArray([
        'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
        'fill' => ['fillType' => 'solid', 'startColor' => ['rgb' => '4472C4']],
        'borders' => ['allBorders' => ['borderStyle' => 'thin']],
        'alignment' => ['horizontal' => 'center'],
    ]);
    $ws->getStyle("D2:D$last")->getNumberFormat()->setFormatCode('#,##0.00');
    $ws->getStyle("H2:H$last")->getNumberFormat()->setFormatCode('0.0000');
    $ws->getColumnDimension('A')->setWidth(24);
    $ws->getColumnDimension('I')->setWidth(32);
    $ws->getRowDimension(1)->setRowHeight(22);
    $ws->freezePane('A2');

    $ws->fromArray([['Customer', 'SKU', 'Qty', 'Price', 'Status', 'Date', 'Total', 'Ratio', 'Note', 'Bucket']]);
    $chunk = [];
    $at = 2;
    foreach (benchRows($rows) as $row) {
        $chunk[] = $row;
        if (\count($chunk) === 2048) {
            $ws->fromArray($chunk, null, 'A' . $at, true);
            $at += 2048;
            $chunk = [];
        }
    }
    if ($chunk) {
        $ws->fromArray($chunk, null, 'A' . $at, true);
    }
    $ws->setAutoFilter("A1:J$last");
    (new $writerClass($spreadsheet))->save($file);
    $spreadsheet->disconnectWorksheets();
}

\gc_collect_cycles();
$t0 = \hrtime(true);
$result = $adapters[$lib][$workload]($file, $rows);
$wall = (\hrtime(true) - $t0) / 1e9;
$peak = \memory_get_peak_usage(true);
$size = \is_file($file) ? \filesize($file) : 0;

\printf("%s,%s,%d,%.3f,%d,%d\n", $lib, $workload, $workload === 'read' ? (int) $result : $rows, $wall, $peak, $size);
