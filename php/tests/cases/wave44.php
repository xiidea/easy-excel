<?php

declare(strict_types=1);

use EasyExcel\Compat\Chart\Chart;
use EasyExcel\Compat\Chart\DataSeries;
use EasyExcel\Compat\Chart\DataSeriesValues;
use EasyExcel\Compat\Chart\Legend;
use EasyExcel\Compat\Chart\PlotArea;
use EasyExcel\Compat\Chart\Title;
use EasyExcel\Compat\RichText\RichText;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\AutoFilter\Column;
use EasyExcel\Compat\Worksheet\AutoFilter\Column\Rule;

return [
    'rich text: runs with fonts serialize to the extension' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $rich = new RichText();
        $rich->createText('Total: ');
        $bold = $rich->createTextRun('42');
        $bold->getFont()->setBold(true);
        $bold->getFont()->getColor()->setRGB('FF0000');
        $ws->setCellValue('A1', $rich);

        $calls = EasyExcelFake::calls('set_rich_text');
        T::same(1, \count($calls));
        T::same('A1', $calls[0][1][2]);
        $runs = $calls[0][1][3];
        T::same('Total: ', $runs[0]['text']);
        T::same(false, isset($runs[0]['font']), 'plain run has no font');
        T::same('42', $runs[1]['text']);
        T::same(true, $runs[1]['font']['bold']);
        T::same('FFFF0000', $runs[1]['font']['color']['argb']);
        // the plain text was buffered so dimensions stay correct
        T::same('Total: 42', $ws->getCell('A1')->getValue());
    },

    'memory drawing: GD resource renders to base64 png' => function (): void {
        if (!\function_exists('imagecreatetruecolor')) {
            return; // ext-gd not present in this runtime
        }
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $img = \imagecreatetruecolor(20, 10);
        $drawing = new \EasyExcel\Compat\Worksheet\MemoryDrawing();
        $drawing->setName('Logo')->setImageResource($img)->setCoordinates('C3')->setWidth(40);
        $drawing->setWorksheet($s->getActiveSheet());

        $calls = EasyExcelFake::calls('add_image_bytes');
        T::same(1, \count($calls));
        T::same('C3', $calls[0][1][2]);
        T::same('.png', $calls[0][1][3]['extension']);
        T::ok($calls[0][1][3]['data'] !== '', 'base64 data present');
        T::same(40, $calls[0][1][3]['width']);
    },

    'chart facade: maps DataSeries onto the native spec' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();

        $labels = [new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$B$1', null, 1)];
        $categories = [new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$A$5', null, 4)];
        $values = [new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$B$2:$B$5', null, 4)];
        $series = new DataSeries(DataSeries::TYPE_BARCHART, DataSeries::GROUPING_STACKED, [0], $labels, $categories, $values);
        $plotArea = new PlotArea(null, [$series]);
        $chart = new Chart('c1', new Title('Sales'), new Legend(Legend::POSITION_BOTTOM), $plotArea, true, 'gap',
            new Title('Month'), new Title('Amount'));
        $chart->setTopLeftPosition('E2');
        $ws->addChart($chart);

        $calls = EasyExcelFake::calls('add_chart');
        T::same(1, \count($calls));
        T::same('E2', $calls[0][1][2]);
        $spec = $calls[0][1][3];
        T::same('colStacked', $spec['type']);
        T::same('Sales', $spec['title']);
        T::same('bottom', $spec['legend']['position']);
        T::same('Month', $spec['xAxisTitle']);
        T::same('Amount', $spec['yAxisTitle']);
        T::same('Worksheet!$B$1', $spec['series'][0]['name']);
        T::same('Worksheet!$A$2:$A$5', $spec['series'][0]['categories']);
        T::same('Worksheet!$B$2:$B$5', $spec['series'][0]['values']);
    },

    'chart facade: bar direction and line type' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $values = [new DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', null, 4)];

        $bar = new DataSeries(DataSeries::TYPE_BARCHART, DataSeries::GROUPING_CLUSTERED, [0], [], [], $values);
        $bar->setPlotDirection(DataSeries::DIRECTION_BAR);
        $ws->addChart((new Chart('b', null, null, new PlotArea(null, [$bar])))->setTopLeftPosition('A1'));

        $line = new DataSeries(DataSeries::TYPE_LINECHART, DataSeries::GROUPING_STANDARD, [0], [], [], $values);
        $ws->addChart((new Chart('l', null, null, new PlotArea(null, [$line])))->setTopLeftPosition('B1'));

        $types = \array_map(static fn (array $c): string => $c[1][3]['type'], EasyExcelFake::calls('add_chart'));
        T::same(['bar', 'line'], $types);
    },

    'auto filter column rules: build excelize expressions' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $filter = $ws->getAutoFilter();
        $filter->setRange('A1:B10');
        $filter->getColumn('A')->createRule()->setRule(Rule::AUTOFILTER_COLUMN_RULE_EQUAL, 'East');
        $filter->getColumn('B')->createRule()->setRule(Rule::AUTOFILTER_COLUMN_RULE_GREATERTHAN, 2000);

        $calls = EasyExcelFake::calls('auto_filter_columns');
        T::ok(\count($calls) >= 1, 'columns applied');
        $last = \end($calls)[1];
        T::same('A1:B10', $last[2]);
        $cols = $last[3];
        $byCol = [];
        foreach ($cols as $c) {
            $byCol[$c['column']] = $c['expression'];
        }
        T::same('x == East', $byCol['A']);
        T::same('x > 2000', $byCol['B']);
    },

    'auto filter column rules: two rules join with or' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $filter = $ws->getAutoFilter();
        $filter->setRange('A1:A10');
        $col = $filter->getColumn('A');
        $col->setJoin(Column::AUTOFILTER_COLUMN_JOIN_OR);
        $col->createRule()->setRule(Rule::AUTOFILTER_COLUMN_RULE_EQUAL, 'East');
        $col->createRule()->setRule(Rule::AUTOFILTER_COLUMN_RULE_EQUAL, 'West');

        $calls = EasyExcelFake::calls('auto_filter_columns');
        $cols = \end($calls)[1][3];
        T::same('x == East or x == West', $cols[0]['expression']);
    },

    'aliases: wave 4.4 classes resolve via bootstrap' => function (): void {
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Chart\\Chart'), 'Chart alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Chart\\DataSeries'), 'DataSeries alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\MemoryDrawing'), 'MemoryDrawing alias');
        T::ok(\class_exists('PhpOffice\\PhpSpreadsheet\\Worksheet\\AutoFilter\\Column'), 'Column alias');
    },
];
