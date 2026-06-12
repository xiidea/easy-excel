<?php

declare(strict_types=1);

use EasyExcel\Compat\Cell\DataValidation;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Style\Conditional;
use EasyExcel\Compat\Worksheet\Drawing;

return [
    'validation: cell-bound setters push the full state' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $v = $s->getActiveSheet()->getCell('C2')->getDataValidation();
        $v->setType(DataValidation::TYPE_LIST)
            ->setFormula1('"open,paid,void"')
            ->setAllowBlank(true)
            ->setShowErrorMessage(true)
            ->setErrorTitle('Bad status');

        $calls = EasyExcelFake::calls('set_validation');
        T::ok(\count($calls) >= 1, 'validation pushed');
        $last = \end($calls)[1];
        T::same('C2', $last[2]);
        T::same('list', $last[3]['type']);
        T::same('"open,paid,void"', $last[3]['formula1']);
        T::same(true, $last[3]['allowBlank']);
        T::same('Bad status', $last[3]['errorTitle']);
    },

    'validation: standalone applied via worksheet range' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $v = new DataValidation();
        $v->setType(DataValidation::TYPE_WHOLE);
        $v->setOperator(DataValidation::OPERATOR_BETWEEN);
        $v->setFormula1('1');
        $v->setFormula2('100');
        T::same([], EasyExcelFake::calls('set_validation'), 'standalone setters must not push');
        $s->getActiveSheet()->setDataValidation('B2:B100', $v);

        $calls = EasyExcelFake::calls('set_validation');
        T::same(1, \count($calls));
        T::same('B2:B100', $calls[0][1][2]);
        T::same('between', $calls[0][1][3]['operator']);
    },

    'conditional: rules with detached styles serialize' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $rule = new Conditional();
        $rule->setConditionType(Conditional::CONDITION_CELLIS)
            ->setOperatorType(Conditional::OPERATOR_GREATERTHAN)
            ->addCondition('5');
        $rule->getStyle()->getFont()->setBold(true);
        $rule->getStyle()->getFill()->setFillType('solid');
        $scale = (new Conditional())->setColorScale('FF0000', '00FF00', 'FFFF00');
        $s->getActiveSheet()->getStyle('B2:B10')->setConditionalStyles([$rule, $scale]);

        T::same([], EasyExcelFake::calls('apply_style'), 'detached styles must not style cells');
        $calls = EasyExcelFake::calls('set_conditional');
        T::same(1, \count($calls));
        [$h, $sheet, $range, $rules] = $calls[0][1];
        T::same('B2:B10', $range);
        T::same(2, \count($rules));
        T::same('cellIs', $rules[0]['type']);
        T::same('greaterThan', $rules[0]['operator']);
        T::same(['5'], $rules[0]['conditions']);
        T::same(['bold' => true], $rules[0]['style']['font']);
        T::same('solid', $rules[0]['style']['fill']['fillType']);
        T::same('colorScale', $rules[1]['type']);
        T::same('FFFF00', $rules[1]['colorScale']['midColor']);
    },

    'drawing: attaches via setWorksheet with scaling info' => function (): void {
        EasyExcelFake::reset();
        $tmp = \tempnam(\sys_get_temp_dir(), 'img');
        \file_put_contents($tmp, 'not-a-real-png');
        $s = new Spreadsheet();
        $d = new Drawing();
        $d->setName('Logo')->setPath($tmp)->setCoordinates('E2')->setOffsetX(5)->setWidth(120);
        T::same([], EasyExcelFake::calls('add_image'), 'configure first, push on attach');
        $d->setWorksheet($s->getActiveSheet());

        $calls = EasyExcelFake::calls('add_image');
        T::same(1, \count($calls));
        T::same('E2', $calls[0][1][2]);
        T::same($tmp, $calls[0][1][3]['path']);
        T::same(120, $calls[0][1][3]['width']);
        T::same(5, $calls[0][1][3]['offsetX']);
        @\unlink($tmp);
    },

    'protection: pushes only once sheet protection is on' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $p = $s->getActiveSheet()->getProtection();
        $p->setPassword('secret');
        T::same([], EasyExcelFake::calls('protect_sheet'), 'no push while sheet=false');
        $p->setSheet(true);
        $p->setFormatCells(true);

        $calls = EasyExcelFake::calls('protect_sheet');
        T::same(2, \count($calls));
        $last = \end($calls)[1][2];
        T::same(true, $last['sheet']);
        T::same('secret', $last['password']);
        T::same(true, $last['formatCells']);
    },

    'chart: native API serializes the spec' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->addNativeChart('E2', [
            'type' => 'col',
            'title' => 'Totals',
            'series' => [['name' => 'A', 'categories' => 'Worksheet!$A$2:$A$5', 'values' => 'Worksheet!$B$2:$B$5']],
            'legend' => ['position' => 'bottom'],
        ]);
        $calls = EasyExcelFake::calls('add_chart');
        T::same(1, \count($calls));
        T::same('E2', $calls[0][1][2]);
        T::same('col', $calls[0][1][3]['type']);
    },

    'toArray: calculateFormulas flag reaches the extension' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', 1);
        $ws->toArray(null, true, true);
        $ws->toArray(null, false, true);

        $reads = EasyExcelFake::calls('read_rows');
        T::same(true, $reads[0][1][5], 'calculateFormulas=true → calc flag set');
        T::same(false, $reads[1][1][5], 'calculateFormulas=false → calc flag clear');
    },

    'aliases: phase-3 classes resolve via bootstrap' => function (): void {
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Cell\\DataValidation'), 'DataValidation alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Style\\Conditional'), 'Conditional alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\Drawing'), 'Drawing alias');
        T::same(
            DataValidation::TYPE_LIST,
            \PhpOffice\PhpSpreadsheet\Cell\DataValidation::TYPE_LIST
        );
    },
];
