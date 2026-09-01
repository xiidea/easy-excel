<?php

declare(strict_types=1);

use EasyExcel\Exception\EasyExcelException;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\Worksheet;

return [
    // ------------------------------------------------------------ page breaks

    'wave53: BREAK_* constants match PhpSpreadsheet' => function (): void {
        T::same(0, Worksheet::BREAK_NONE);
        T::same(1, Worksheet::BREAK_ROW);
        T::same(2, Worksheet::BREAK_COLUMN);
        T::same(16383, Worksheet::BREAK_ROW_MAX_COLUMN);
    },

    'wave53: setBreak sends a row break to the extension' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        T::ok($ws->setBreak('A24', Worksheet::BREAK_ROW) === $ws, 'fluent');

        $calls = EasyExcelFake::calls('set_break');
        T::same(1, \count($calls));
        T::same('A24', $calls[0][1][2], 'cell');
        T::same(Worksheet::BREAK_ROW, $calls[0][1][3], 'break type');
    },

    'wave53: setBreak sends a column break' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setBreak('O1', Worksheet::BREAK_COLUMN);

        $calls = EasyExcelFake::calls('set_break');
        T::same(Worksheet::BREAK_COLUMN, $calls[0][1][3]);
    },

    'wave53: setBreak defaults to BREAK_NONE, which removes' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setBreak('A24');

        $calls = EasyExcelFake::calls('set_break');
        T::same(Worksheet::BREAK_NONE, $calls[0][1][3], 'default removes rather than adds');
    },

    'wave53: setBreak accepts a [column, row] array coordinate' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setBreak([15, 24], Worksheet::BREAK_ROW);

        $calls = EasyExcelFake::calls('set_break');
        T::same('O24', $calls[0][1][2], 'column 15 is O');
    },

    'wave53: setBreakByColumnAndRow maps to the same call' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setBreakByColumnAndRow(3, 10, Worksheet::BREAK_ROW);

        $calls = EasyExcelFake::calls('set_break');
        T::same('C10', $calls[0][1][2]);
        T::same(Worksheet::BREAK_ROW, $calls[0][1][3]);
    },

    'wave53: an invalid break type surfaces the extension error' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        T::throws(
            EasyExcelException::class,
            static fn () => $s->getActiveSheet()->setBreak('A1', 99),
            'unknown break type must fail loudly',
        );
    },

    'wave53: an invalid break cell surfaces the extension error' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        T::throws(
            EasyExcelException::class,
            static fn () => $s->getActiveSheet()->setBreak('not-a-cell', Worksheet::BREAK_ROW),
            'bad cell reference must fail loudly',
        );
    },

    'wave53: multiple breaks accumulate' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setBreak('A24', Worksheet::BREAK_ROW);
        $ws->setBreak('A48', Worksheet::BREAK_ROW);
        $ws->setBreak('O1', Worksheet::BREAK_COLUMN);

        T::same(3, \count(EasyExcelFake::calls('set_break')), 'each break is its own op');
    },

    // -------------------------------------------------------------- selection

    'wave53: setSelectedCells sends the range' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        T::ok($ws->setSelectedCells('B2:D5') === $ws, 'fluent');

        $calls = EasyExcelFake::calls('set_selection');
        T::same(1, \count($calls));
        T::same('B2:D5', $calls[0][1][2]);
    },

    'wave53: setSelectedCells accepts a single cell' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setSelectedCells('C3');

        T::same('C3', EasyExcelFake::calls('set_selection')[0][1][2]);
    },

    'wave53: setSelectedCells accepts a [column, row] array' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setSelectedCells([2, 7]);

        T::same('B7', EasyExcelFake::calls('set_selection')[0][1][2]);
    },

    'wave53: setSelectedCell and ByColumnAndRow are aliases' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->setSelectedCell('D4');
        $ws->setSelectedCellByColumnAndRow(5, 9);

        $calls = EasyExcelFake::calls('set_selection');
        T::same('D4', $calls[0][1][2]);
        T::same('E9', $calls[1][1][2]);
    },

    'wave53: an invalid selection surfaces the extension error' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        T::throws(
            EasyExcelException::class,
            static fn () => $s->getActiveSheet()->setSelectedCells('nope!'),
            'bad selection must fail loudly',
        );
    },

    'wave53: selection after freezePane issues both calls' => function (): void {
        // The Go side merges them onto one pane record; from PHP both ops must
        // still be sent, in order, so the merge has something to merge.
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->freezePane('A2');
        $ws->setSelectedCells('B5');

        T::same(1, \count(EasyExcelFake::calls('freeze_panes')));
        T::same(1, \count(EasyExcelFake::calls('set_selection')));
    },

    'wave53: extension errors are catchable as the broad Compat exception' => function (): void {
        // Native::check throws EasyExcelException, which Compat\Exception
        // extends — so a broad `catch (PhpSpreadsheet\Exception)` does NOT
        // catch these. Pinned because it is the opposite of what it looks like.
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $caught = null;
        try {
            $s->getActiveSheet()->setBreak('A1', 99);
        } catch (\RuntimeException $e) {
            $caught = $e;
        }
        T::ok($caught instanceof EasyExcelException, 'is the native exception');
        T::ok(!($caught instanceof \EasyExcel\Compat\Exception), 'is NOT the Compat subclass');
    },

    // --------------------------------------------------- calculateColumnWidths

    'wave53: calculateColumnWidths is an accepted no-op' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $before = \count(EasyExcelFake::$log);
        T::ok($ws->calculateColumnWidths() === $ws, 'returns $this for chaining');
        T::same($before, \count(EasyExcelFake::$log), 'touches the extension not at all');
    },

    'wave53: calculateColumnWidths chains with real work' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->getColumnDimension('A')->setAutoSize(true);
        $ws->calculateColumnWidths()->setCellValue('A1', 'wide enough to matter');
        $ws->flush();

        T::same('wide enough to matter', $ws->getCell('A1')->getValue(), 'chaining works');
    },

    // --------------------------------------------------------------- aliasing

    'wave53: the new methods exist on the PhpOffice alias' => function (): void {
        $cls = 'PhpOffice\PhpSpreadsheet\Worksheet\Worksheet';
        T::ok(\class_exists($cls), 'alias resolves');
        foreach (['setBreak', 'setSelectedCells', 'calculateColumnWidths'] as $m) {
            T::ok(\method_exists($cls, $m), "$m exists on the alias");
        }
        T::same(1, \constant($cls . '::BREAK_ROW'), 'constants come through the alias');
    },
];
