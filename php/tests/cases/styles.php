<?php

declare(strict_types=1);

use EasyExcel\Compat\Cell\Hyperlink;
use EasyExcel\Compat\NamedRange;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Style\Border;
use EasyExcel\Compat\Style\Color;
use EasyExcel\Compat\Style\Fill;
use EasyExcel\Compat\Worksheet\PageSetup;

return [
    'style: font setters send partial specs' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $style = $s->getActiveSheet()->getStyle('A1:C1');
        $style->getFont()->setBold(true)->setSize(14.5);

        $calls = EasyExcelFake::calls('apply_style');
        T::same(2, \count($calls), 'one CGO call per setter');
        T::same('A1:C1', $calls[0][1][2]);
        T::same(['font' => ['bold' => true]], $calls[0][1][3]);
        T::same(['font' => ['size' => 14.5]], $calls[1][1][3]);
        T::same(true, $style->getFont()->getBold());
        T::same(14.5, $style->getFont()->getSize());
    },

    'style: applyFromArray passes the whole array through' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $array = [
            'font' => ['bold' => true, 'color' => ['rgb' => 'FF0000']],
            'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['argb' => 'FFFFFF00']],
            'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
            'alignment' => ['horizontal' => 'center', 'wrapText' => true],
            'numberFormat' => ['formatCode' => '0.00'],
        ];
        $s->getActiveSheet()->getStyle('B2:D4')->applyFromArray($array);

        $calls = EasyExcelFake::calls('apply_style');
        T::same(1, \count($calls));
        T::same($array, $calls[0][1][3]);
    },

    'style: color proxies notify their component' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $style = $s->getActiveSheet()->getStyle('A1');
        $style->getFont()->getColor()->setRGB('00FF00');
        $style->getFill()->setFillType(Fill::FILL_SOLID);
        $style->getFill()->getStartColor()->setARGB('FFABCDEF');
        $style->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_MEDIUM);
        $style->getBorders()->getTop()->getColor()->setRGB('123456');

        $specs = \array_map(static fn (array $c): array => $c[1][3], EasyExcelFake::calls('apply_style'));
        T::same(['font' => ['color' => ['argb' => 'FF00FF00']]], $specs[0]);
        T::same(['fill' => ['fillType' => 'solid']], $specs[1]);
        T::same(['fill' => ['startColor' => ['argb' => 'FFABCDEF']]], $specs[2]);
        T::same(['borders' => ['allBorders' => ['borderStyle' => 'medium']]], $specs[3]);
        T::same(['borders' => ['top' => ['color' => ['argb' => 'FF123456']]]], $specs[4]);
    },

    'style: standalone Color helpers' => function (): void {
        $c = new Color(Color::COLOR_RED);
        T::same('FFFF0000', $c->getARGB());
        T::same('FF0000', $c->getRGB());
        $c->setRGB('336699');
        T::same('FF336699', $c->getARGB());
    },

    'dimensions: column width, autosize, row height' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->getColumnDimension('B')->setWidth(25.5);
        $ws->getColumnDimensionByColumn(3)->setAutoSize(true);
        $ws->getRowDimension(1)->setRowHeight(30);

        T::same([1, 'Worksheet', 2, 2, 25.5], EasyExcelFake::calls('set_col_width')[0][1]);
        T::same([1, 'Worksheet', 3, 3], EasyExcelFake::calls('set_col_autosize')[0][1]);
        T::same([1, 'Worksheet', 1, 30.0], EasyExcelFake::calls('set_row_height')[0][1]);
    },

    'structure: auto-filter, freeze panes, unfreeze' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setAutoFilter('A1:F100');
        $ws->freezePane('B2');
        $ws->unfreezePane();

        T::same('A1:F100', EasyExcelFake::calls('auto_filter')[0][1][2]);
        $freeze = EasyExcelFake::calls('freeze_panes');
        T::same('B2', $freeze[0][1][2]);
        T::same('', $freeze[1][1][2]);
    },

    'hyperlink: via cell and via worksheet' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->getCell('A1')->getHyperlink()->setUrl('https://example.com');
        $ws->setHyperlink('B2', new Hyperlink('https://example.org', 'docs'));

        $calls = EasyExcelFake::calls('set_hyperlink');
        T::same([1, 'Worksheet', 'A1', 'https://example.com', ''], $calls[0][1]);
        T::same([1, 'Worksheet', 'B2', 'https://example.org', 'docs'], $calls[1][1]);
    },

    'comment: text runs accumulate and replace' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $comment = $s->getActiveSheet()->getComment('C3');
        $comment->setAuthor('QA');
        $comment->getText()->createTextRun('first ');
        $comment->getText()->createTextRun('second');

        $calls = EasyExcelFake::calls('set_comment');
        T::same([1, 'Worksheet', 'C3', 'QA', 'first '], $calls[0][1]);
        T::same([1, 'Worksheet', 'C3', 'QA', 'first second'], $calls[1][1]);
        T::same('first second', $comment->getText()->getPlainText());
    },

    'named range: sheet-qualified refersTo' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $s->addNamedRange(new NamedRange('data', $ws, '$A$1:$C$10'));
        $s->addNamedRange(new NamedRange('local', $ws, 'B2', true));

        $calls = EasyExcelFake::calls('defined_name');
        T::same([1, 'data', 'Worksheet!$A$1:$C$10', ''], $calls[0][1]);
        T::same([1, 'local', 'Worksheet!B2', 'Worksheet'], $calls[1][1]);
    },

    'page setup: setters accumulate state' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $setup = $s->getActiveSheet()->getPageSetup();
        $setup->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
        $setup->setPaperSize(PageSetup::PAPERSIZE_A4);
        $setup->setFitToWidth(1);

        $calls = EasyExcelFake::calls('page_setup');
        T::same([1, 'Worksheet', 'landscape', -1, -1, -1], $calls[0][1]);
        T::same([1, 'Worksheet', 'landscape', 9, -1, -1], $calls[1][1]);
        T::same([1, 'Worksheet', 'landscape', 9, 1, -1], $calls[2][1]);
    },

    'merge + number format: no buffer flush (stream-friendly order)' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setCellValue('A1', 'Header');                  // buffered
        $ws->mergeCells('A1:C1');                           // must not flush
        $ws->getStyle('A1')->getNumberFormat()->setFormatCode('0.00');
        T::same([], EasyExcelFake::calls('write_rows'), 'style/merge ops must not flush the row buffer');
        $ws->flush();
        T::same(1, \count(EasyExcelFake::calls('write_rows')));
        T::same('A1:C1', EasyExcelFake::calls('merge_cells')[0][1][2]);
        T::same('0.00', EasyExcelFake::calls('set_number_format')[0][1][3]);
    },

    'richtext: run formatting raises a clear unsupported error' => function (): void {
        $s = new Spreadsheet();
        $run = $s->getActiveSheet()->getComment('A1')->getText()->createTextRun('x');
        T::throws(\EasyExcel\Compat\Exception::class, static fn () => $run->getFont());
    },

    'aliases: phase-2 classes resolve via bootstrap' => function (): void {
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Style\\Fill'), 'Fill alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Style\\Border'), 'Border alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\PageSetup'), 'PageSetup alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\NamedRange'), 'NamedRange alias');
        T::same(Fill::FILL_SOLID, \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
    },
];
