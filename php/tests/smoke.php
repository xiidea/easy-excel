<?php

declare(strict_types=1);

/*
 * End-to-end smoke test against the REAL extension (not the fake ABI).
 * Run inside the built image:
 *
 *   docker run --rm -v $PWD/php:/opt/easy-excel/php frankenphp-easy-excel \
 *       frankenphp php-cli /opt/easy-excel/php/tests/smoke.php
 *
 * Exercises the full stack: alias bootstrap -> shim buffering -> CGO bridge
 * -> Go core -> excelize -> a real .xlsx on disk, then reads it back.
 */

$root = \dirname(__DIR__);
\spl_autoload_register(static function (string $class) use ($root): void {
    if (\str_starts_with($class, 'EasyExcel\\')) {
        $path = $root . '/src/EasyExcel/' . \str_replace('\\', '/', \substr($class, 10)) . '.php';
        if (\is_file($path)) {
            require $path;
        }
    }
});
require $root . '/src/bootstrap.php';

function check(bool $cond, string $what): void
{
    if (!$cond) {
        \fwrite(STDERR, "SMOKE FAIL: $what\n");
        exit(1);
    }
    echo "  ok  $what\n";
}

check(\function_exists('easy_excel_new'), 'extension is loaded (easy_excel_new exists)');
echo '  ..  extension version: ' . \easy_excel_version() . "\n";

// --- write path -------------------------------------------------------------
$s = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$ws = $s->getActiveSheet();
check($ws->getTitle() === 'Worksheet', 'default sheet name');

$rows = [];
for ($i = 1; $i <= 10000; ++$i) {
    $rows[] = ["customer-$i", $i, \round($i * 1.37, 2), $i % 2 === 0 ? 'paid' : 'open'];
}
$t0 = \hrtime(true);
$ws->fromArray($rows);
$ws->setCellValue('F1', '=B2+B3');
$ws->setCellValue('F2', '0042');
$ws->setCellValueExplicit('F3', '=not_formula', \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
$ws->setCellValue('F4', new DateTime('2026-06-11 12:00:00', new DateTimeZone('UTC')));
$ws->getStyle('F4')->getNumberFormat()->setFormatCode('yyyy-mm-dd');

$file = \sys_get_temp_dir() . '/smoke.xlsx';
\PhpOffice\PhpSpreadsheet\IOFactory::createWriter($s, 'Xlsx')->save($file);
$ms = (int) ((\hrtime(true) - $t0) / 1e6);
check(\is_file($file) && \filesize($file) > 10000, "10k rows written + saved ({$ms}ms, " . \filesize($file) . ' bytes)');
$s->disconnectWorksheets();

// --- read path ---------------------------------------------------------------
$r = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
$rs = $r->getActiveSheet();
check($rs->getCell('A1')->getValue() === 'customer-1', 'string round-trip');
check($rs->getCell('B2')->getValue() === 2.0, 'number round-trip');
check($rs->getCell('F2')->getValue() === '0042', 'leading-zero string preserved');
check($rs->getCell('F1')->getValue() === '=B2+B3', 'formula source preserved');
$calc = $rs->getCell('F1')->getCalculatedValue();
check($calc == 5, 'formula calculated by excelize engine (got ' . \var_export($calc, true) . ')');
check($rs->getCell('F3')->getValue() === '=not_formula', 'explicit string not a formula');
check($rs->getCell('F4')->getFormattedValue() === '2026-06-11', 'date number format applied');
check($rs->getHighestRow() === 10000, 'highest row');
$sum = 0;
foreach ($rs->toArray(null, true, false) as $row) {
    $sum += \count($row);
}
check($sum >= 40000, "toArray chunked read ($sum cells)");
$r->disconnectWorksheets();
@\unlink($file);

// --- phase 2: styles & structure end-to-end --------------------------------------
$s3 = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$ws3 = $s3->getActiveSheet();

// the report pattern: style first, then bulk-write — must stay streaming
$ws3->getStyle('A1:C1')->applyFromArray([
    'font' => ['bold' => true, 'color' => ['rgb' => 'FFFFFF']],
    'fill' => ['fillType' => 'solid', 'startColor' => ['rgb' => '4472C4']],
    'borders' => ['allBorders' => ['borderStyle' => 'thin']],
    'alignment' => ['horizontal' => 'center'],
]);
$ws3->getStyle('B2:B5000')->getNumberFormat()->setFormatCode('#,##0.00');
$ws3->getColumnDimension('A')->setWidth(28);
$ws3->getRowDimension(1)->setRowHeight(24);
$ws3->freezePane('A2');
$ws3->mergeCells('E1:F1');

$data = [['Name', 'Amount', 'Status']];
for ($i = 2; $i <= 5000; ++$i) {
    $data[] = ["item-$i", $i * 1.5, 'ok'];
}
$ws3->fromArray($data);

// save-time deferred ops (documented one-time degrade)
$ws3->setAutoFilter('A1:C5000');
$ws3->getCell('A2')->getHyperlink()->setUrl('https://example.com/item-2');
$ws3->getComment('B2')->getText()->createTextRun('reviewed');
$ws3->getColumnDimension('C')->setAutoSize(true);
$ws3->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
$s3->addNamedRange(new \PhpOffice\PhpSpreadsheet\NamedRange('amounts', $ws3, '$B$2:$B$5000'));

$styled = \sys_get_temp_dir() . '/smoke-styled.xlsx';
$t1 = \hrtime(true);
\PhpOffice\PhpSpreadsheet\IOFactory::createWriter($s3, 'Xlsx')->save($styled);
$ms = (int) ((\hrtime(true) - $t1) / 1e6);
check(\is_file($styled) && \filesize($styled) > 5000, "styled report saved ({$ms}ms, " . \filesize($styled) . ' bytes)');
$s3->disconnectWorksheets();

$r3 = \PhpOffice\PhpSpreadsheet\IOFactory::load($styled);
$rs3 = $r3->getActiveSheet();
check($rs3->getCell('A1')->getValue() === 'Name', 'styled file header intact');
check($rs3->getCell('B3')->getFormattedValue() === '4.50', 'streamed number format renders');
check($rs3->getHighestRow() === 5000, 'styled file row count');
$r3->disconnectWorksheets();
@\unlink($styled);

// --- csv + load control surface ------------------------------------------------
$s2 = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$s2->getActiveSheet()->fromArray([['a', '-x', 'b,c']]);
$csv = \sys_get_temp_dir() . '/smoke.csv';
$w = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($s2, 'Csv');
$w->setSanitizeFormulas(true);
$w->save($csv);
$content = (string) \file_get_contents($csv);
check(\str_contains($content, "'-x") && \str_contains($content, '"b,c"'), 'csv guard + quoting');
$s2->disconnectWorksheets();
@\unlink($csv);

$stats = \easy_excel_stats();
check($stats[0] === 0, 'no leaked native handles');
check($stats[1] === 0, 'memory accounting back to zero');

echo "\nSMOKE PASS\n";
