<?php

declare(strict_types=1);

use EasyExcel\Compat\Chart\Axis;
use EasyExcel\Compat\Chart\GridLines;
use EasyExcel\Compat\Reader\Csv as CsvReader;
use EasyExcel\Compat\Reader\IReader;
use EasyExcel\Compat\Reader\Xlsx as XlsxReader;
use EasyExcel\Compat\Shared\StringHelper;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\Worksheet;

/**
 * Closes the method-level gaps left by the wave suites: accessors and
 * behaviours that exist on classes this PR adds but were never asserted.
 */
return [
    // ---- Chart\Axis gridline accessors -----------------------------------

    'coverage: axis gridline getters round-trip and default to null' => function (): void {
        $axis = new Axis();
        T::same(null, $axis->getMajorGridlines(), 'no major gridlines by default');
        T::same(null, $axis->getMinorGridlines(), 'no minor gridlines by default');

        $major = new GridLines();
        $minor = new GridLines();
        T::ok($axis->setMajorGridlines($major) === $axis, 'fluent');
        T::ok($axis->getMajorGridlines() === $major, 'major returns the same instance');
        T::same(null, $axis->getMinorGridlines(), 'setting major leaves minor alone');

        $axis->setMinorGridlines($minor);
        T::ok($axis->getMinorGridlines() === $minor, 'minor returns the same instance');
    },

    'coverage: gridlines can be cleared with null' => function (): void {
        $axis = new Axis();
        $axis->setMajorGridlines(new GridLines());
        T::same(true, $axis->buildSpec()['majorGridlines'], 'set');
        $axis->setMajorGridlines(null);
        T::ok(!\array_key_exists('majorGridlines', $axis->buildSpec()), 'cleared');
    },

    // ---- GridLines::activate ---------------------------------------------

    'coverage: GridLines::activate is fluent and keeps the object on' => function (): void {
        $g = new GridLines();
        T::same(true, $g->getObjectState(), 'constructed active');
        T::ok($g->activate() === $g, 'fluent');
        T::same(true, $g->getObjectState(), 'still active');
    },

    // ---- Reader\IReader::canRead -----------------------------------------

    'coverage: canRead accepts a real xlsx and rejects other bytes' => function (): void {
        $xlsx = \tempnam(\sys_get_temp_dir(), 'cov') . '.xlsx';
        (new \EasyExcel\Compat\Writer\Xlsx(new Spreadsheet()))->save($xlsx);

        $reader = new XlsxReader();
        T::same(true, $reader->canRead($xlsx), 'a saved workbook is readable');

        $notXlsx = \tempnam(\sys_get_temp_dir(), 'cov') . '.txt';
        \file_put_contents($notXlsx, "just text, no zip header\n");
        T::same(false, $reader->canRead($notXlsx), 'plain text is not xlsx');
        T::same(false, $reader->canRead('/nonexistent/easy-excel/nope.xlsx'), 'missing file');

        \unlink($xlsx);
        \unlink($notXlsx);
    },

    'coverage: canRead is reachable through the IReader contract' => function (): void {
        // The point of the interface: consumers type-hint it and call canRead
        // without knowing which reader they hold.
        $probe = static fn (IReader $r, string $f): bool => $r->canRead($f);

        $csv = \tempnam(\sys_get_temp_dir(), 'cov') . '.csv';
        \file_put_contents($csv, "a,b\n1,2\n");
        T::same(true, $probe(new CsvReader(), $csv), 'csv reader via the interface');
        \unlink($csv);
    },

    // ---- Shared\StringHelper::strToTitle ---------------------------------

    'coverage: strToTitle is UTF-8 aware' => function (): void {
        T::same('Hello World', StringHelper::strToTitle('hello world'));
        T::same('Élan Vital', StringHelper::strToTitle('élan vital'), 'accented initial upper-cases');
        T::same('', StringHelper::strToTitle(''), 'empty is safe');
    },

    // ---- Worksheet: BREAK_* behaviours not previously asserted -----------

    'coverage: setBreak rejects a bad type before touching the extension' => function (): void {
        EasyExcelFake::reset();
        $ws = (new Spreadsheet())->getActiveSheet();
        T::throws(
            \EasyExcel\Exception\EasyExcelException::class,
            static fn () => $ws->setBreak('A1', 42),
            'an unknown break constant must fail loudly',
        );
    },

    'coverage: BREAK_ROW_MAX_COLUMN is the documented constant' => function (): void {
        // Carried for source compatibility; assert it rather than assume.
        T::same(16383, Worksheet::BREAK_ROW_MAX_COLUMN);
    },

    // ---- Chart\Layout geometry setters -----------------------------------

    'coverage: Layout geometry setters round-trip' => function (): void {
        // The existing wave-5.4 test only reads geometry supplied through the
        // constructor, so the setters were never exercised.
        $l = new \EasyExcel\Compat\Chart\Layout();
        T::same(null, $l->getXPosition(), 'unset by default');

        T::ok($l->setXPosition(0.25) === $l, 'fluent');
        T::same(0.25, $l->getXPosition());
        $l->setYPosition(0.5);
        T::same(0.5, $l->getYPosition());
        $l->setWidth(0.75);
        T::same(0.75, $l->getWidth());
        $l->setHeight(0.4);
        T::same(0.4, $l->getHeight());

        T::same('', $l->getLayoutTarget(), 'empty by default');
        $l->setLayoutTarget('inner');
        T::same('inner', $l->getLayoutTarget());

        // Nullable dimensions can be cleared again.
        $l->setWidth(null);
        T::same(null, $l->getWidth());
    },

    // ---- Writer\Pdf\Snappy accessor --------------------------------------

    'coverage: the Snappy instance round-trips through the accessor' => function (): void {
        EasyExcelFake::reset();
        $w = new \EasyExcel\Compat\Writer\Pdf\Snappy(new Spreadsheet());
        T::same(null, $w->getSnappy(), 'none injected by default');

        // Stands in for knplabs/knp-snappy, which is not a dependency here.
        $snappy = new class {
            public function getOutputFromHtml(string $html, array $options = []): string
            {
                return '%PDF-1.4';
            }
        };
        T::ok($w->setSnappy($snappy) === $w, 'fluent');
        T::ok($w->getSnappy() === $snappy, 'returns the same instance');
    },

    // ---- Settings: the accepted no-op contract ---------------------------

    'coverage: setChartRenderer survives a reset cycle' => function (): void {
        \EasyExcel\Compat\Settings::reset();
        \EasyExcel\Compat\Settings::setChartRenderer('Some\\Renderer');
        \EasyExcel\Compat\Settings::unsetChartRenderer();
        T::same(null, \EasyExcel\Compat\Settings::getChartRenderer(), 'unset clears');
        \EasyExcel\Compat\Settings::reset();
    },
];
