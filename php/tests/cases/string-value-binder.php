<?php

declare(strict_types=1);

/*
 * PhpSpreadsheet's Cell\StringValueBinder — everything binds as a string
 * unless conversion for that value type is suppressed.
 */

use EasyExcel\Compat\Cell\StringValueBinder;
use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;

return [
    'string binder: numerics and booleans bind as strings by default' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->setValueBinder(new StringValueBinder());
        $sheet = $s->getActiveSheet();
        $sheet->setCellValue('A1', 123.45);
        $sheet->setCellValue('B1', true);
        $sheet->setCellValue('C1', '=SUM(A1:B1)');
        $sheet->fromArray([[42, 'text']], null, 'A2');
        $sheet->flush();

        T::same('123.45', $sheet->getCell('A1')->getValue());
        T::same('1', $sheet->getCell('B1')->getValue());
        T::same('=SUM(A1:B1)', $sheet->getCell('C1')->getValue());
        T::same('42', $sheet->getCell('A2')->getValue());
        T::same('text', $sheet->getCell('B2')->getValue());
    },

    'string binder: suppressed conversions keep native types' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $binder = (new StringValueBinder())
            ->setNumericConversion(false)
            ->setFormulaConversion(false);
        $s->setValueBinder($binder);
        $sheet = $s->getActiveSheet();
        $sheet->setCellValue('A1', 123.45);
        $sheet->setCellValue('B1', '=SUM(1,2)');
        $sheet->setCellValue('C1', true);
        $sheet->flush();

        T::same(123.45, $sheet->getCell('A1')->getValue());
        T::same('=SUM(1,2)', $sheet->getCell('B1')->getValue());
        T::same('1', $sheet->getCell('C1')->getValue(), 'boolean conversion still on -> string');
    },

    'string binder: DateTime binds as formatted string, unstringable throws' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->setValueBinder(new StringValueBinder());
        $sheet = $s->getActiveSheet();
        $sheet->setCellValue('A1', new \DateTime('2025-07-01 08:30:00'));
        $sheet->flush();

        T::same('2025-07-01 08:30:00', $sheet->getCell('A1')->getValue());

        T::throws(Exception::class, static function () use ($sheet): void {
            $sheet->setCellValue('B1', [1, 2, 3]);
        }, 'arrays are unstringable');
    },

    'string binder: aliases as PhpOffice\\PhpSpreadsheet\\Cell\\StringValueBinder' => function (): void {
        EasyExcelFake::reset();
        T::ok(
            \class_exists('PhpOffice\\PhpSpreadsheet\\Cell\\StringValueBinder'),
            'StringValueBinder must be part of the aliased Compat surface'
        );
    },
];
