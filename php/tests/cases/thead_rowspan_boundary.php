<?php

declare(strict_types=1);

/*
 * A third HTML-rendering regression from the same PDF diff: a header whose
 * merge straddles the repeat-rows boundary.
 *
 * setRowsToRepeatAtTopByStartAndEnd(9, 10) puts rows 9-10 in <thead>, but a
 * report whose column headers merge rows 10-11 (one row of captions, one row
 * of sub-captions) has its merges start on the last thead row and end on the
 * first tbody row. A rowspan cannot cross that section boundary — the browser
 * clamps it at the end of <thead> — so the covered cells never reserved their
 * columns and the sub-caption row slid left under the wrong headers.
 */

return [
    'thead: a merge straddling the boundary extends the header section' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->fromArray([
            ['Account Code', 'Account Name', 'Opening Balance', 'For the Period', 'For the Period', 'Closing Balance'],
            ['', '', '', 'Dr.', 'Cr.', ''],
        ], null, 'A10');
        $worksheet->fromArray([['1101010101-01', 'Cash in Hand', '1.00', '2.00', '3.00', '4.00']], null, 'A12');
        foreach (['A', 'B', 'C', 'F'] as $column) {
            $worksheet->mergeCells($column . '10:' . $column . '11');
        }
        $worksheet->mergeCells('D10:E10');
        $worksheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(9, 10);

        $html = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll();

        T::ok((bool) preg_match('#<thead>(.*?)</thead>#s', $html, $m), 'a thead is emitted');
        T::same(3, substr_count($m[1], '<tr'), 'thead must stretch to cover the row-11 half of the merge');
        T::ok(str_contains($m[1], 'Dr.'), 'the sub-caption row belongs to the header section');
        T::ok(str_contains($m[1], 'Cr.'), 'the sub-caption row belongs to the header section');
    },

    'thead: a merge wholly inside the repeat rows leaves the boundary alone' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->fromArray([['Title', '', '', '', '', '']], null, 'A9');
        $worksheet->fromArray([['A', 'B', 'C', 'D', 'E', 'F']], null, 'A10');
        $worksheet->fromArray([['1', '2', '3', '4', '5', '6']], null, 'A11');
        $worksheet->mergeCells('A9:F9');
        $worksheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(9, 10);

        $html = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll();

        preg_match('#<thead>(.*?)</thead>#s', $html, $m);
        T::same(2, substr_count($m[1] ?? '', '<tr'), 'the header section keeps its configured size');
    },

    'thead: a merge starting after the repeat rows leaves the boundary alone' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->fromArray([['A', 'B']], null, 'A9');
        $worksheet->fromArray([['C', 'D']], null, 'A10');
        $worksheet->fromArray([['E', 'F']], null, 'A11');
        $worksheet->fromArray([['G', 'H']], null, 'A12');
        $worksheet->mergeCells('A11:A12');
        $worksheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(9, 10);

        $html = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll();

        preg_match('#<thead>(.*?)</thead>#s', $html, $m);
        T::same(2, substr_count($m[1] ?? '', '<tr'), 'a body-only merge must not pull rows into the header');
    },
];
