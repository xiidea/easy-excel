<?php

declare(strict_types=1);

use EasyExcel\Compat\Cell\DataType;
use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;

return [
    'spreadsheet: default sheet is "Worksheet"' => function (): void {
        $s = new Spreadsheet();
        T::same('Worksheet', $s->getActiveSheet()->getTitle());
        T::same(1, $s->getSheetCount());
        T::same(0, $s->getActiveSheetIndex());
    },

    'worksheet: value binding through buffer' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', 'Hello')
            ->setCellValue('B1', '123')          // numeric string -> number
            ->setCellValue('C1', '0123')         // leading zero -> stays string
            ->setCellValue('D1', '=SUM(A1:B1)')  // formula
            ->setCellValue('E1', true);
        T::same('Hello', $ws->getCell('A1')->getValue());
        T::same(123.0, $ws->getCell('B1')->getValue());
        T::same('0123', $ws->getCell('C1')->getValue());
        T::same('=SUM(A1:B1)', $ws->getCell('D1')->getValue());
        T::same(true, $ws->getCell('E1')->getValue());
        T::same(DataType::TYPE_FORMULA, $ws->getCell('D1')->getDataType());
    },

    'worksheet: explicit types defeat auto-binding' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValueExplicit('A1', '=NOT_A_FORMULA()', DataType::TYPE_STRING);
        $ws->setCellValueExplicit('B1', '19.5', DataType::TYPE_NUMERIC);
        T::same('=NOT_A_FORMULA()', $ws->getCell('A1')->getValue());
        T::same(19.5, $ws->getCell('B1')->getValue());
    },

    'worksheet: DateTime values become excel serials' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', new DateTime('2023-06-15 12:00:00', new DateTimeZone('UTC')));
        T::same(45092.5, $ws->getCell('A1')->getValue());
    },

    'worksheet: fromArray + toArray round-trip' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->fromArray([
            ['name', 'qty'],
            ['widget', 3],
            ['gadget', 5],
        ], null, 'B2');
        $out = $ws->toArray(null, true, false);
        // toArray always starts at A1: row 1 and column A are null padding
        T::same(4, \count($out));
        T::same([null, null, null], $out[0]);
        T::same([null, 'name', 'qty'], $out[1]);
        T::same([null, 'widget', 3], $out[2]);
        T::same(4, $ws->getHighestRow());
        T::same('C', $ws->getHighestColumn());
    },

    'worksheet: rangeToArray with cell refs' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->fromArray([['a', 'b'], ['c', 'd']], null, 'A1');
        $out = $ws->rangeToArray('A1:B2', null, true, false, true);
        T::same(['A' => 'a', 'B' => 'b'], $out[1]);
        T::same(['A' => 'c', 'B' => 'd'], $out[2]);
    },

    'worksheet: write-behind buffer batches rows' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        for ($r = 1; $r <= 600; ++$r) {
            $ws->setCellValue("A$r", $r);
        }
        // threshold is 512: one automatic flush must have happened already
        $auto = EasyExcelFake::calls('write_rows');
        T::same(1, \count($auto), 'exactly one flush at the 512-row threshold');
        T::same(512, $auto[0][1][4], 'flush carried 512 rows');
        T::same(1, $auto[0][1][2], 'anchored at row 1');
        // remaining 88 rows arrive on read
        T::same(600.0, $ws->getCell('A600')->getValue());
        T::same(2, \count(EasyExcelFake::calls('write_rows')));
    },

    'worksheet: sparse buffer flushes as contiguous runs' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A5', 'five');
        $ws->setCellValue('A2', 'two');   // out of order in the buffer
        $ws->setCellValue('B2', 'btwo');
        $ws->flush();
        $calls = EasyExcelFake::calls('write_rows');
        T::same(2, \count($calls), 'two runs: row 2 and row 5');
        T::same(2, $calls[0][1][2], 'first run anchored at row 2 (sorted)');
        T::same(5, $calls[1][1][2]);
        T::same('two', $ws->getCell('A2')->getValue());
        T::same('five', $ws->getCell('A5')->getValue());
    },

    'worksheet: same-cell rewrites coalesce in the buffer' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', 'first');
        $ws->setCellValue('A1', 'second');
        $ws->flush();
        T::same(1, \count(EasyExcelFake::calls('write_rows')), 'one run, no boundary chatter');
        T::same('second', $ws->getCell('A1')->getValue());
    },

    'spreadsheet: sheet management' => function (): void {
        $s = new Spreadsheet();
        $data = $s->createSheet();
        T::same(2, $s->getSheetCount());
        $data->setTitle('Data');
        T::same('Data', $s->getSheetNames()[1]);
        T::same($data, $s->getSheetByName('Data'));
        T::same(1, $s->getIndex($data));
        $s->setActiveSheetIndex(1);
        T::same('Data', $s->getActiveSheet()->getTitle());
        $s->setActiveSheetIndex(0);
        $s->removeSheetByIndex(1);
        T::same(1, $s->getSheetCount());
        T::same(null, $s->getSheetByName('Data'));
    },

    'spreadsheet: disconnect releases the native handle' => function (): void {
        $s = new Spreadsheet();
        T::same(1, \count(EasyExcelFake::$store));
        $s->disconnectWorksheets();
        T::same(0, \count(EasyExcelFake::$store), 'native workbook freed');
        T::throws(Exception::class, static fn () => $s->getActiveSheet(), 'disconnected use throws');
    },

    'style: number format reaches the extension' => function (): void {
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', 0.125);
        $ws->getStyle('A1')->getNumberFormat()->setFormatCode('0.00%');
        $calls = EasyExcelFake::calls('set_number_format');
        T::same(1, \count($calls));
        T::same(['A1', '0.00%'], [\end($calls)[1][2], \end($calls)[1][3]]);
    },

    'style: phase-2 components route through apply_style' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $style = $s->getActiveSheet()->getStyle('A1');
        $style->applyFromArray(['font' => ['bold' => true]]);
        $style->getFont()->setItalic(true);
        T::same(2, \count(EasyExcelFake::calls('apply_style')));
    },
];
