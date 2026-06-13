<?php

declare(strict_types=1);

use EasyExcel\Compat\Spreadsheet;

return [
    'rows: insert and remove shift data' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->fromArray([['r1'], ['r2'], ['r3']]);
        $ws->insertNewRowBefore(2, 2);

        T::same('r1', $ws->getCell('A1')->getValue());
        T::same(null, $ws->getCell('A2')->getValue());
        T::same('r2', $ws->getCell('A4')->getValue());
        T::same(5, $ws->getHighestRow());

        $ws->removeRow(2, 2);
        T::same('r2', $ws->getCell('A2')->getValue());
        T::same(3, $ws->getHighestRow());
    },

    'columns: insert and remove by letter and index' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->fromArray([['a', 'b', 'c']]);
        $ws->insertNewColumnBefore('B');

        T::same('b', $ws->getCell('C1')->getValue());
        T::same(null, $ws->getCell('B1')->getValue());

        $ws->removeColumnByIndex(2);
        T::same('b', $ws->getCell('B1')->getValue());
        T::same([1, 'Worksheet', 2, 1], EasyExcelFake::calls('remove_cols')[0][1]);
    },

    'sheets: createSheet at index and copySheet' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->createSheet();             // Worksheet1, appended
        $first = $s->createSheet(0);   // inserted at the front

        T::same(0, $s->getIndex($first));
        T::same([1, $first->getTitle(), 0], EasyExcelFake::calls('move_sheet')[0][1]);

        $s->getSheetByName('Worksheet')->setCellValue('A1', 'src');
        $s->getSheetByName('Worksheet')->flush();
        $copy = $s->copySheet('Worksheet', 'Backup');
        T::same('Backup', $copy->getTitle());
        T::same('src', $copy->getCell('A1')->getValue());
    },

    'sheet view: gridlines, zoom, tab color accumulate' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setShowGridlines(false);
        $ws->getSheetView()->setZoomScale(75);
        $ws->getTabColor()->setRGB('FF0000');

        $calls = EasyExcelFake::calls('sheet_view');
        T::same(3, \count($calls));
        $last = \end($calls)[1][2];
        T::same(false, $last['showGridlines']);
        T::same(75, $last['zoomScale']);
        T::same('FF0000', $last['tabColor']);
    },

    'header/footer: placeholder codes pass through' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $hf = $s->getActiveSheet()->getHeaderFooter();
        $hf->setOddHeader('&C&BReport');
        $hf->setOddFooter('&CPage &P of &N');

        $last = \end(EasyExcelFake::$log)[1][2];
        T::same('&C&BReport', $last['oddHeader']);
        T::same('&CPage &P of &N', $last['oddFooter']);
        T::same('&C&BReport', $hf->getOddHeader());
    },

    'page margins: setters accumulate inches' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $margins = $s->getActiveSheet()->getPageMargins();
        $margins->setTop(1.25)->setLeft(0.7);

        $last = \end(EasyExcelFake::$log)[1][2];
        T::same(1.25, $last['top']);
        T::same(0.7, $last['left']);
        T::same(-1, $last['bottom']);
        T::same(1.25, $margins->getTop());
    },

    'aliases: wave 4.3 classes resolve via bootstrap' => function (): void {
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\SheetView'), 'SheetView alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\HeaderFooter'), 'HeaderFooter alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\PageMargins'), 'PageMargins alias');
    },
];
