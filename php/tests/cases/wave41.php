<?php

declare(strict_types=1);

use EasyExcel\Compat\Calculation\Calculation;
use EasyExcel\Compat\Cell\Cell;
use EasyExcel\Compat\Cell\DataType;
use EasyExcel\Compat\Cell\DefaultValueBinder;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Style\Borders;
use EasyExcel\Compat\Style\Conditional;
use EasyExcel\Compat\Style\Fill;
use EasyExcel\Compat\Writer\Xlsx as XlsxWriter;

/** binds integers above 2^53 as strings — the ERP probe's binder */
final class BigIdBinder extends DefaultValueBinder
{
    public function bindValue(Cell $cell, mixed $value): bool
    {
        if (\is_numeric($value) && !\str_contains((string) $value, '.') && $value > 9007199254740992) {
            $cell->setValueExplicit((string) $value, DataType::TYPE_STRING);

            return true;
        }

        return parent::bindValue($cell, $value);
    }
}

/** restores the no-custom-binder state Cell has at process start */
function resetBinder(): void
{
    $prop = new ReflectionProperty(Cell::class, 'valueBinder');
    $prop->setValue(null, null);
}

return [
    'binder: custom binder intercepts setCellValue and fromArray' => function (): void {
        EasyExcelFake::reset();
        resetBinder();
        Cell::setValueBinder(new BigIdBinder());
        try {
            $s = new Spreadsheet();
            $ws = $s->getActiveSheet();
            $ws->setCellValue('A1', '9007199254740995'); // > 2^53 → string
            $ws->setCellValue('B1', '123');              // normal binding rules
            $ws->fromArray([['9007199254741002', 7]], null, 'A2', true);
            $ws->flush();

            T::same('9007199254740995', $ws->getCell('A1')->getValue(), 'big ID stays a string');
            T::same(123.0, $ws->getCell('B1')->getValue(), 'default rules still apply');
            T::same('9007199254741002', $ws->getCell('A2')->getValue(), 'fromArray routed through binder');
            T::same(7.0, $ws->getCell('B2')->getValue());
        } finally {
            resetBinder();
        }
    },

    'binder: DefaultValueBinder::dataTypeForValue parity' => function (): void {
        T::same(DataType::TYPE_NULL, DefaultValueBinder::dataTypeForValue(null));
        T::same(DataType::TYPE_BOOL, DefaultValueBinder::dataTypeForValue(true));
        T::same(DataType::TYPE_NUMERIC, DefaultValueBinder::dataTypeForValue('123'));
        T::same(DataType::TYPE_STRING, DefaultValueBinder::dataTypeForValue('0123'));
        T::same(DataType::TYPE_FORMULA, DefaultValueBinder::dataTypeForValue('=SUM(A1)'));
        T::same(DataType::TYPE_STRING, DefaultValueBinder::dataTypeForValue('hello'));
    },

    'binder: no custom binder keeps the fromArray fast path' => function (): void {
        EasyExcelFake::reset();
        resetBinder();
        $s = new Spreadsheet();
        $s->getActiveSheet()->fromArray([[1, 2], [3, 4]]);
        $writes = EasyExcelFake::calls('write_rows');
        T::same(1, \count($writes), 'one bulk write, not per-cell buffering');
        T::same(2, $writes[0][1][4], 'both rows in one batch');
    },

    'properties: setters accumulate and push' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getProperties()->setTitle('ERP Report')->setSubject('Accounts')->setCompany('BRAC');

        $calls = EasyExcelFake::calls('doc_props');
        T::same(3, \count($calls));
        $last = \end($calls)[1][1];
        T::same('ERP Report', $last['title']);
        T::same('Accounts', $last['subject']);
        T::same('BRAC', $last['company']);
        T::same('ERP Report', $s->getProperties()->getTitle());
    },

    'properties: custom properties + created/modified' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $props = $s->getProperties();
        $props->setCustomProperty('Reviewed', true);
        $props->setCustomProperty('Revision', 3, \EasyExcel\Compat\Document\Properties::PROPERTY_TYPE_INTEGER);
        $props->setCustomProperty('Ratio', 1.5);
        $props->setCreated('2026-06-13 10:00:00');

        T::same(true, $props->isCustomPropertySet('Reviewed'));
        T::same(3, $props->getCustomPropertyValue('Revision'));
        T::same('i', $props->getCustomPropertyType('Revision'));
        T::same('f', $props->getCustomPropertyType('Ratio'));
        T::same(['Reviewed', 'Revision', 'Ratio'], $props->getCustomProperties());

        $custom = EasyExcelFake::calls('custom_prop');
        T::same('Reviewed', $custom[0][1][1]['name']);
        T::same('b', $custom[0][1][1]['type']);
        T::same(true, $custom[0][1][1]['value']);

        $props->removeCustomProperty('Reviewed');
        T::same(false, $props->isCustomPropertySet('Reviewed'));
        $custom = EasyExcelFake::calls('custom_prop');
        $last = \end($custom)[1][1];
        T::same(true, $last['remove']);

        // created reaches doc_props as an ISO timestamp
        $docCalls = EasyExcelFake::calls('doc_props');
        $doc = \end($docCalls)[1][1];
        T::ok(\str_starts_with($doc['created'] ?? '', '2026-06-13T10:00:00'), 'created timestamp');
    },

    'print titles and print area become reserved defined names' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $setup = $s->getActiveSheet()->getPageSetup();
        $setup->setRowsToRepeatAtTopByStartAndEnd(6, 7);
        $setup->setPrintArea('A1:G15');

        $calls = EasyExcelFake::calls('defined_name');
        T::same(['_xlnm.Print_Titles', 'Worksheet!$6:$7', 'Worksheet'],
            [$calls[0][1][1], $calls[0][1][2], $calls[0][1][3]]);
        T::same(['_xlnm.Print_Area', 'Worksheet!$A$1:$G$15', 'Worksheet'],
            [$calls[1][1][1], $calls[1][1][2], $calls[1][1][3]]);
    },

    'conditional styles: getter returns session rules' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $rule = (new Conditional())->setConditionType(Conditional::CONDITION_CELLIS)
            ->setOperatorType(Conditional::OPERATOR_GREATERTHAN)->addCondition('5');
        $ws->getStyle('B2:B9')->setConditionalStyles([$rule]);

        $got = $ws->getStyle('B2:B9')->getConditionalStyles();
        T::same(1, \count($got));
        T::ok($got[0] === $rule, 'same rule instance returned');
        T::same([], $ws->getStyle('C1')->getConditionalStyles(), 'other ranges empty');
    },

    'encryption: password flows through writer and reader' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setCellValue('A1', 'x');
        $tmp = \tempnam(\sys_get_temp_dir(), 'enc');
        (new XlsxWriter($s))->setPassword('s3cret')->save($tmp);
        T::same('s3cret', \end(EasyExcelFake::$log)[1][2] ?? EasyExcelFake::calls('save_xlsx')[0][1][2]);

        try {
            (new \EasyExcel\Compat\Reader\Xlsx())->setPassword('s3cret')->load($tmp);
        } catch (\Throwable) {
            // the fake cannot open files; the call log is what matters
        }
        $open = EasyExcelFake::calls('open');
        T::same('s3cret', $open[0][1][1], 'reader password reaches the ABI');
        @\unlink($tmp);
    },

    'gradient fill: rotation and colors reach the style spec' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $fill = $s->getActiveSheet()->getStyle('A1')->getFill();
        $fill->setFillType(Fill::FILL_GRADIENT_LINEAR);
        $fill->setRotation(90.5);
        $fill->getStartColor()->setRGB('FF0000');
        $fill->getEndColor()->setRGB('0000FF');

        $specs = \array_map(static fn (array $c): array => $c[1][3], EasyExcelFake::calls('apply_style'));
        T::same(['fill' => ['fillType' => 'linear']], $specs[0]);
        T::same(['fill' => ['rotation' => 90.5]], $specs[1]);
        T::same(['fill' => ['startColor' => ['argb' => 'FFFF0000']]], $specs[2]);
        T::same(['fill' => ['endColor' => ['argb' => 'FF0000FF']]], $specs[3]);
    },

    'diagonal borders: direction + style reach the spec' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $borders = $s->getActiveSheet()->getStyle('B2')->getBorders();
        $borders->getDiagonal()->setBorderStyle(\EasyExcel\Compat\Style\Border::BORDER_THIN);
        $borders->setDiagonalDirection(Borders::DIAGONAL_BOTH);

        $specs = \array_map(static fn (array $c): array => $c[1][3], EasyExcelFake::calls('apply_style'));
        T::same(['borders' => ['diagonal' => ['borderStyle' => 'thin']]], $specs[0]);
        T::same(['borders' => ['diagonalDirection' => 3]], $specs[1]);
        T::same(Borders::DIAGONAL_BOTH, $borders->getDiagonalDirection());
    },

    'unmerge + merge getter' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->mergeCells('A1:C1');
        $ws->mergeCells('A2:C2');
        $ws->unmergeCells('A1:C1');

        T::same('A1:C1', EasyExcelFake::calls('unmerge_cells')[0][1][2]);
        T::same(['A2:C2' => 'A2:C2'], $ws->getMergeCells());
    },

    'calculation: cache controls are tolerated no-ops' => function (): void {
        $calc = Calculation::getInstance(new Spreadsheet());
        $calc->disableCalculationCache();
        T::same(false, $calc->getCalculationCacheEnabled());
        $calc->enableCalculationCache();
        T::same(true, $calc->getCalculationCacheEnabled());
        $calc->clearCalculationCache();
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Calculation\\Calculation'), 'alias resolves');
    },

    'aliases: wave 4.1 classes resolve via bootstrap' => function (): void {
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Cell\\DefaultValueBinder'), 'DefaultValueBinder alias');
        T::ok(\interface_exists('PhpOffice\\PhpSpreadsheet\\Cell\\IValueBinder'), 'IValueBinder alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Document\\Properties'), 'Properties alias');
    },
];
