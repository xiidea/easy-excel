<?php

declare(strict_types=1);

use EasyExcel\Compat\Settings;
use EasyExcel\Compat\Shared\File as SharedFile;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Writer\Xlsx;

return [
    // ---- review item 3: precalc must be opt-in ---------------------------

    'review3: an untouched writer does not enable precalculation' => function (): void {
        EasyExcelFake::reset();
        (new Xlsx(new Spreadsheet()))->save(\tempnam(\sys_get_temp_dir(), 'rf') . '.xlsx');
        $calls = EasyExcelFake::calls('set_precalculate_formulas');
        T::same(false, $calls[0][1][1], 'inheriting upstream\'s default must not cost streaming');
    },

    'review3: getPreCalculateFormulas still reports upstream default' => function (): void {
        T::same(true, (new Xlsx(new Spreadsheet()))->getPreCalculateFormulas(), 'API parity kept');
    },

    'review3: DISABLE_PRECALCULATE_FORMULAE stays off' => function (): void {
        EasyExcelFake::reset();
        $w = new Xlsx(new Spreadsheet());
        $w->setPreCalculateFormulas(true);
        $w->save(\tempnam(\sys_get_temp_dir(), 'rf') . '.xlsx', Xlsx::DISABLE_PRECALCULATE_FORMULAE);
        $calls = EasyExcelFake::calls('set_precalculate_formulas');
        T::same(false, $calls[0][1][1], 'the save flag overrides an explicit opt-in');
    },

    // ---- review item 5: dimension caches ---------------------------------

    'review5: column dimensions are case-insensitive' => function (): void {
        EasyExcelFake::reset();
        $ws = (new Spreadsheet())->getActiveSheet();
        $ws->getColumnDimension('a')->setWidth(40);
        T::same(40.0, $ws->getColumnDimension('A')->getWidth(), "'a' and 'A' are one column");
        $ws->getColumnDimension('A')->setWidth(15);
        T::same(15.0, $ws->getColumnDimension('a')->getWidth(), 'and stay in sync both ways');
    },

    'review5: the row dimension cache is bounded' => function (): void {
        EasyExcelFake::reset();
        $ws = (new Spreadsheet())->getActiveSheet();
        // Past the cap a fresh instance is returned rather than retained, so a
        // per-row loop over a large export cannot grow without bound.
        for ($r = 1; $r <= 5000; ++$r) {
            $ws->getRowDimension($r)->setRowHeight(20);
        }
        $ref = new \ReflectionProperty($ws, 'rowDimensions');
        $ref->setAccessible(true);
        T::ok(\count($ref->getValue($ws)) <= 4096, 'cache capped');
    },

    'review5: rows within the cap still read back' => function (): void {
        EasyExcelFake::reset();
        $ws = (new Spreadsheet())->getActiveSheet();
        $ws->getRowDimension(3)->setRowHeight(33);
        T::same(33.0, $ws->getRowDimension(3)->getRowHeight(), 'read-back preserved');
    },

    // ---- getCalculatedValue must not mangle numeric-looking strings ------

    'calc: a formatted string result keeps its leading zeros' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        // is_numeric('0042') is true, so a bare cast would return 42 and drop
        // the zero padding the formula exists to produce.
        EasyExcelFake::$calculated['A1'] = '0042';
        T::same('0042', $ws->getCell('A1')->getCalculatedValue(), 'string preserved');
    },

    'calc: the string/number split matches upstream PhpSpreadsheet' => function (): void {
        // Verified against phpoffice/phpspreadsheet 5.8: TEXT(A1,"0000")
        // returns the string '0042' and ="1.50" the string '1.50', with
        // floats only for genuine numeric results. Casting every is_numeric()
        // string moved *away* from upstream, not toward it.
        EasyExcelFake::reset();
        $ws = (new Spreadsheet())->getActiveSheet();
        foreach (['0042', '1.50', ' 3', '1e3'] as $i => $raw) {
            $cell = 'A' . ($i + 1);
            EasyExcelFake::$calculated[$cell] = $raw;
            T::same($raw, $ws->getCell($cell)->getCalculatedValue(), "$raw stays a string");
        }
    },

    'calc: a genuine numeric result is still cast to a number' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        EasyExcelFake::$calculated['A1'] = '1542';
        T::same(1542, $ws->getCell('A1')->getCalculatedValue(), 'canonical integer casts');
        EasyExcelFake::$calculated['A2'] = '-0.75';
        T::same(-0.75, $ws->getCell('A2')->getCalculatedValue(), 'canonical float casts');
    },

    // ---- review item 6: Settings matches the real upstream surface -------

    'review6: Settings exposes upstream members, not invented ones' => function (): void {
        foreach (['setLocale', 'getLocale', 'htmlEntityFlags', 'setChartRenderer', 'setCache', 'setHttpClient'] as $m) {
            T::ok(\method_exists(Settings::class, $m), "$m exists upstream");
        }
        foreach (['setLibXmlLoaderOptions', 'setLibXmlDisableEntityLoader'] as $m) {
            T::ok(!\method_exists(Settings::class, $m), "$m is not an upstream method");
        }
    },

    // ---- review item 8: upload temp dir flag ------------------------------

    'review8: the upload-temp-dir flag survives an empty ini' => function (): void {
        SharedFile::reset();
        SharedFile::setUseUploadTempDirectory(true);
        T::same(true, SharedFile::getUseUploadTempDirectory(), 'flag stored, not the resolved path');
        T::ok(\is_dir(SharedFile::sysGetTempDir()), 'still resolves to a real directory');
        SharedFile::reset();
        T::same(false, SharedFile::getUseUploadTempDirectory());
    },
];
