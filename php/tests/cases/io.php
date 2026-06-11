<?php

declare(strict_types=1);

use EasyExcel\Compat\IOFactory;
use EasyExcel\Compat\Reader\Csv as CsvReader;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Writer\Csv as CsvWriter;
use EasyExcel\Compat\Writer\Xlsx as XlsxWriter;

return [
    'iofactory: identify and create' => function (): void {
        T::same('Xlsx', IOFactory::identify('/tmp/report.XLSX'));
        T::same('Csv', IOFactory::identify('data.csv'));
        $s = new Spreadsheet();
        T::ok(IOFactory::createWriter($s, 'Xlsx') instanceof XlsxWriter, 'xlsx writer');
        T::ok(IOFactory::createWriter($s, 'Csv') instanceof CsvWriter, 'csv writer');
        T::throws(\EasyExcel\Compat\Exception::class, static fn () => IOFactory::createWriter($s, 'Ods'));
    },

    'writer: xlsx save flushes buffers and writes the file' => function (): void {
        $s = new Spreadsheet();
        $s->getActiveSheet()->setCellValue('A1', 'persisted');
        $file = \tempnam(\sys_get_temp_dir(), 'eex') . '.xlsx';
        try {
            (new XlsxWriter($s))->save($file);
            T::ok(\is_file($file) && \filesize($file) > 0, 'file written');
            T::same(1, \count(EasyExcelFake::calls('save_xlsx')));
            // the buffered A1 must have been flushed before saving
            T::ok(\str_contains((string) \file_get_contents($file), 'persisted'), 'buffer flushed before save');
        } finally {
            @\unlink($file);
        }
    },

    'writer: csv content, delimiter and guard' => function (): void {
        $s = new Spreadsheet();
        $s->getActiveSheet()->fromArray([
            ['a', 'b;c', '-danger'],
            ['x', 'say "hi"', '2'],
        ]);
        $file = \tempnam(\sys_get_temp_dir(), 'eex') . '.csv';
        try {
            (new CsvWriter($s))->setDelimiter(';')->setSanitizeFormulas(true)->save($file);
            $content = (string) \file_get_contents($file);
            T::ok(\str_contains($content, '"b;c"'), 'delimiter collision quoted');
            T::ok(\str_contains($content, "'-danger"), 'injection guard applied');
            T::ok(\str_contains($content, '"say ""hi"""'), 'quote escaping');
        } finally {
            @\unlink($file);
        }
    },

    'writer: php:// stream target works via temp file' => function (): void {
        $s = new Spreadsheet();
        $s->getActiveSheet()->setCellValue('A1', 'streamed');
        \ob_start();
        (new XlsxWriter($s))->save('php://output');
        $out = \ob_get_clean();
        T::ok(\str_contains($out, 'streamed'), 'content reached the stream');
    },

    'reader: csv loads in chunks with binding' => function (): void {
        $file = \tempnam(\sys_get_temp_dir(), 'eex') . '.csv';
        \file_put_contents($file, "name,qty\nwidget,3\ngadget,0150\n");
        try {
            $s = (new CsvReader())->load($file);
            $ws = $s->getActiveSheet();
            T::same('name', $ws->getCell('A1')->getValue());
            T::same(3.0, $ws->getCell('B2')->getValue(), 'numeric string bound to number');
            T::same('0150', $ws->getCell('B3')->getValue(), 'leading zero preserved');
        } finally {
            @\unlink($file);
        }
    },

    'bootstrap: PhpOffice\\PhpSpreadsheet aliases resolve' => function (): void {
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Spreadsheet'), 'root class aliased');
        $s = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        T::ok($s instanceof Spreadsheet, 'alias is the compat class');
        $s->getActiveSheet()->setCellValue('A1', 'via alias');
        T::same('via alias', $s->getActiveSheet()->getCell('A1')->getValue());
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Cell\\Coordinate'), 'nested namespace aliased');
        T::same(27, \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString('AA'));
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($s, 'Xlsx');
        T::ok($writer instanceof XlsxWriter, 'IOFactory through alias');
    },
];
