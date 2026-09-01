<?php

declare(strict_types=1);

/*
 * Two HTML-rendering regressions found while diffing a Compat-rendered PDF
 * against the same report under real phpoffice/phpspreadsheet:
 *
 *  1. getColumnDimension()/getRowDimension() handed out a fresh object per
 *     call, so a width written through one call was unreadable through the
 *     next. Widths still reached the engine (xlsx was fine), but the Html
 *     writer reads them back via getWidth() and so fell back to its default
 *     cell width for every column.
 *
 *  2. The Html writer skipped borders whose style is 'none' and ignored the
 *     'allBorders' key entirely. An explicit BORDER_NONE therefore emitted no
 *     CSS at all and could not override the sheet-wide `table.sheet td`
 *     border, while setAllBorders() emitted nothing whatsoever.
 */

return [
    'column dimension: a width written is readable back' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->getColumnDimension('A')->setWidth(18);

        T::same(18.0, $worksheet->getColumnDimension('A')->getWidth(), 'width must survive a re-fetch');
    },

    'column dimension: the same instance is returned per column' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        T::ok(
            $worksheet->getColumnDimension('A') === $worksheet->getColumnDimension('A'),
            'getColumnDimension must be idempotent'
        );
        T::ok(
            $worksheet->getColumnDimension('A') !== $worksheet->getColumnDimension('B'),
            'distinct columns keep distinct dimensions'
        );
    },

    'column dimension: byColumn shares the cache with byName' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->getColumnDimensionByColumn(1)->setWidth(24);

        T::same(24.0, $worksheet->getColumnDimension('A')->getWidth());
    },

    'row dimension: a height written is readable back' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->getRowDimension(3)->setRowHeight(24);

        T::same(24.0, $worksheet->getRowDimension(3)->getRowHeight());
    },

    'html: column widths reach the generated markup' => function (): void {
        $cellCss = static function (string $html): string {
            preg_match_all('/table\\.sheet \\.(\\w+) \\{([^}]*)\\}/', $html, $m, PREG_SET_ORDER);

            return implode(' | ', array_map(static fn (array $x): string => trim($x[2]), $m));
        };
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->setCellValue('A1', 'x');
        $worksheet->getColumnDimension('A')->setWidth(18);

        $html = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll();

        T::ok(str_contains($html, 'width: 144px'), 'an 18-unit column is 144px, not the default width');
    },

    'html: an explicit BORDER_NONE emits a rule that beats the sheet default' => function (): void {
        $cellCss = static function (string $html): string {
            preg_match_all('/table\\.sheet \\.(\\w+) \\{([^}]*)\\}/', $html, $m, PREG_SET_ORDER);

            return implode(' | ', array_map(static fn (array $x): string => trim($x[2]), $m));
        };
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->setCellValue('A1', 'x');
        $worksheet->getStyle('A1')->getBorders()->getAllBorders()
            ->setBorderStyle(PhpOffice\PhpSpreadsheet\Style\Border::BORDER_NONE);

        $css = $cellCss(PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll());

        T::same(4, substr_count($css, ': none;'), 'all four sides must be cleared');
    },

    'html: an unstyled cell keeps the sheet-wide border' => function (): void {
        $cellCss = static function (string $html): string {
            preg_match_all('/table\\.sheet \\.(\\w+) \\{([^}]*)\\}/', $html, $m, PREG_SET_ORDER);

            return implode(' | ', array_map(static fn (array $x): string => trim($x[2]), $m));
        };
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'x');

        $html = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll();

        T::ok(str_contains($html, 'table.sheet td { border: 1px solid #d0d0d0'), 'sheet default stays');
        T::ok(!str_contains($cellCss($html), 'border'), 'no per-cell border rule is emitted');
    },

    'html: allBorders renders instead of being ignored' => function (): void {
        $cellCss = static function (string $html): string {
            preg_match_all('/table\\.sheet \\.(\\w+) \\{([^}]*)\\}/', $html, $m, PREG_SET_ORDER);

            return implode(' | ', array_map(static fn (array $x): string => trim($x[2]), $m));
        };
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->setCellValue('A1', 'x');
        $worksheet->getStyle('A1')->getBorders()->getAllBorders()
            ->setBorderStyle(PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

        $css = $cellCss(PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll());

        foreach (['top', 'bottom', 'left', 'right'] as $side) {
            T::ok(str_contains($css, "border-{$side}: 1px solid"), "{$side} border must render");
        }
    },

    'html: a single-side border does not leak to the other sides' => function (): void {
        $cellCss = static function (string $html): string {
            preg_match_all('/table\\.sheet \\.(\\w+) \\{([^}]*)\\}/', $html, $m, PREG_SET_ORDER);

            return implode(' | ', array_map(static fn (array $x): string => trim($x[2]), $m));
        };
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->setCellValue('A1', 'x');
        $worksheet->getStyle('A1')->getBorders()->getTop()
            ->setBorderStyle(PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);

        $css = $cellCss(PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Html')->generateHtmlAll());

        T::ok(str_contains($css, 'border-top: 1px solid'), 'the set side renders');
        T::ok(!str_contains($css, 'border-bottom: 1px solid'), 'the unset sides stay unset');
    },
];
