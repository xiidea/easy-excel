<?php

declare(strict_types=1);

use EasyExcel\Compat\Cell\Cell;
use EasyExcel\Compat\Cell\DataType;
use EasyExcel\Compat\Cell\DefaultValueBinder;
use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\Worksheet;

/** binds everything as string — mirrors the ERP's report StringValueBinder */
final class AllStringBinder extends DefaultValueBinder
{
    public function bindValue(Cell $cell, mixed $value): bool
    {
        if (\is_scalar($value)) {
            $cell->setValueExplicit((string) $value, DataType::TYPE_STRING);

            return true;
        }

        return parent::bindValue($cell, $value);
    }
}

return [
    'wave45: Spreadsheet::setValueBinder routes setCellValue and fromArray' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        T::same(null, $s->getValueBinder(), 'no workbook binder by default');
        T::ok($s->setValueBinder(new AllStringBinder()) === $s, 'fluent');

        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', 123);
        $ws->fromArray([[456, 7.5]], null, 'A2', true);
        $ws->flush();

        T::same('123', $ws->getCell('A1')->getValue(), 'setCellValue routed through workbook binder');
        T::same('456', $ws->getCell('A2')->getValue(), 'fromArray routed through workbook binder');
        T::same('7.5', $ws->getCell('B2')->getValue());

        $s->setValueBinder(null);
        $ws->setCellValue('A3', 123);
        $ws->flush();
        T::same(123.0, $ws->getCell('A3')->getValue(), 'null restores default binding');
    },

    'wave45: workbook binder wins over legacy static Cell binder' => function (): void {
        EasyExcelFake::reset();
        $prop = new ReflectionProperty(Cell::class, 'valueBinder');
        $prop->setValue(null, null);
        Cell::setValueBinder(new DefaultValueBinder()); // static set, default rules
        try {
            $s = new Spreadsheet();
            $s->setValueBinder(new AllStringBinder());
            $ws = $s->getActiveSheet();
            $ws->setCellValue('A1', 99);
            $ws->flush();
            T::same('99', $ws->getCell('A1')->getValue(), 'instance binder overrides static');
        } finally {
            $prop->setValue(null, null);
        }
    },

    'wave45: detached Worksheet construct, title, attach via addSheet' => function (): void {
        EasyExcelFake::reset();
        $ws = new Worksheet();
        T::same('Worksheet', $ws->getTitle(), 'default title');
        T::same(null, $ws->getParent(), 'detached');
        $ws->setTitle('Region North');
        T::same('Region North', $ws->getTitle(), 'title set locally while detached');
        T::same(0, \count(EasyExcelFake::calls('rename_sheet')), 'no native rename while detached');

        $s = new Spreadsheet();
        $returned = $s->addSheet($ws);
        T::ok($returned === $ws, 'addSheet returns the sheet');
        T::ok($ws->getParent() === $s, 'attached');
        T::same(2, $s->getSheetCount());
        T::ok($s->getSheetByName('Region North') === $ws, 'find by title');

        $ws->setCellValue('A1', 'hello');
        $ws->flush();
        T::same('hello', $ws->getCell('A1')->getValue(), 'writable after attach');
    },

    'wave45: addSheet honours the index and duplicate titles throw' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = new Worksheet(null, 'First');
        $s->addSheet($ws, 0);
        T::same(0, $s->getIndex($ws), 'inserted at requested index');
        T::same(1, \count(EasyExcelFake::calls('move_sheet')), 'native move issued');

        T::throws(Exception::class, static fn () => $s->addSheet(new Worksheet(null, 'First')));

        $other = new Spreadsheet();
        T::throws(Exception::class, static fn () => $other->addSheet($ws), 'cross-workbook rebind rejected');
    },

    'wave45: detached sheet operations fail loudly' => function (): void {
        EasyExcelFake::reset();
        $ws = new Worksheet(null, 'Loose');
        T::throws(Exception::class, static fn () => $ws->setCellValue('A1', 1)->flush());
    },

    'wave45: eager aliasing covers instanceof and typed parameters' => function (): void {
        // instanceof/type checks never trigger autoload, so these only pass
        // when the alias exists up front (eagerAliasCompat at bootstrap).
        EasyExcelFake::reset();
        $ws = (new Spreadsheet())->getActiveSheet();
        T::ok($ws instanceof \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet, 'instanceof aliased name');
        $typed = static fn (\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet): string => $sheet->getTitle();
        T::same('Worksheet', $typed($ws), 'typed parameter accepts Compat object');
    },
];
