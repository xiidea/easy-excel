<?php

declare(strict_types=1);

use EasyExcel\Compat\Chart\Axis;
use EasyExcel\Compat\Chart\Chart;
use EasyExcel\Compat\Chart\ChartColor;
use EasyExcel\Compat\Chart\DataSeries;
use EasyExcel\Compat\Chart\DataSeriesValues;
use EasyExcel\Compat\Chart\GridLines;
use EasyExcel\Compat\Chart\Layout;
use EasyExcel\Compat\Chart\PlotArea;
use EasyExcel\Compat\Chart\Title;
use EasyExcel\Compat\Spreadsheet;

/** Builds the minimal series graph every chart test needs. */
function w54series(): DataSeries
{
    return new DataSeries(
        DataSeries::TYPE_BARCHART,
        DataSeries::GROUPING_CLUSTERED,
        [0],
        [new DataSeriesValues('String', 'Sheet1!$A$1', null, 1)],
        [new DataSeriesValues('String', 'Sheet1!$B$1:$B$3', null, 3)],
        [new DataSeriesValues('Number', 'Sheet1!$C$1:$C$3', null, 3)],
    );
}

function w54spec(Chart $chart): array
{
    return $chart->buildSpec();
}

return [
    // ------------------------------------------------------------------- Axis

    'wave54: setAxisOptionsProperties binds positionally like PhpSpreadsheet' => function (): void {
        $axis = new Axis();
        // The overwhelmingly common call in the audited app.
        $axis->setAxisOptionsProperties('none');
        T::same('none', $axis->getAxisOptionsProperty('axis_labels'));
        T::same(['labels' => 'none'], $axis->buildSpec());
    },

    'wave54: axis bounds and units reach the spec' => function (): void {
        $axis = new Axis();
        $axis->setAxisOptionsProperties(
            Axis::AXIS_LABELS_LOW,
            null,
            null,
            null,
            null,
            null,
            0,      // minimum
            100,    // maximum
            25,     // majorUnit
        );
        $spec = $axis->buildSpec();
        T::same('low', $spec['labels']);
        T::same(0.0, $spec['minimum'], 'an explicit zero minimum survives');
        T::same(100.0, $spec['maximum']);
        T::same(25.0, $spec['majorUnit']);
    },

    'wave54: unset axis bounds are omitted, not sent as zero' => function (): void {
        $axis = new Axis();
        $axis->setAxisOptionsProperties('low');
        $spec = $axis->buildSpec();
        T::ok(!\array_key_exists('minimum', $spec), 'no minimum key');
        T::ok(!\array_key_exists('maximum', $spec), 'no maximum key');
        T::ok(!\array_key_exists('majorUnit', $spec), 'no majorUnit key');
    },

    'wave54: maxMin orientation becomes reverseOrder' => function (): void {
        $axis = new Axis();
        $axis->setAxisOptionsProperties('low', null, null, Axis::AXIS_ORIENTATION_MAX_MIN);
        T::same(true, $axis->buildSpec()['reverseOrder']);

        $normal = new Axis();
        $normal->setAxisOptionsProperties('low', null, null, Axis::AXIS_ORIENTATION_MIN_MAX);
        T::ok(!\array_key_exists('reverseOrder', $normal->buildSpec()), 'minMax is the default');
    },

    'wave54: log base reaches the spec' => function (): void {
        $axis = new Axis();
        $axis->setAxisOptionsProperties(
            'low', null, null, null, null, null, null, null, null, null, null,
            null, null, null, null,
            10, // logBase
        );
        T::same(10.0, $axis->buildSpec()['logBase']);
    },

    'wave54: axis number format reaches the spec' => function (): void {
        $axis = new Axis();
        $axis->setAxisNumberProperties('0.00%');
        T::same('0.00%', $axis->getAxisNumberFormat());
        T::same('0.00%', $axis->buildSpec()['numFmt']);
    },

    'wave54: an untouched axis contributes nothing' => function (): void {
        T::same([], (new Axis())->buildSpec(), 'no axis block for a default axis');
    },

    'wave54: setFillParameters round-trips through ChartColor' => function (): void {
        $axis = new Axis();
        $axis->setFillParameters('#ffffff');
        T::same('FFFFFF', $axis->getFillProperty('value'), 'normalised: hash stripped, upper-cased');
    },

    // -------------------------------------------------------------- GridLines

    'wave54: gridlines attached to an axis set the spec flags' => function (): void {
        $axis = new Axis();
        $axis->setMajorGridlines(new GridLines());
        T::same(true, $axis->buildSpec()['majorGridlines']);
        T::ok(!\array_key_exists('minorGridlines', $axis->buildSpec()), 'minor stays off');

        $axis->setMinorGridlines(new GridLines());
        T::same(true, $axis->buildSpec()['minorGridlines']);
    },

    'wave54: setLineColorProperties is accepted and round-trips' => function (): void {
        $g = new GridLines();
        T::ok($g->setLineColorProperties('ffffff', 100, GridLines::EXCEL_COLOR_TYPE_ARGB) === $g, 'fluent');
        T::same('FFFFFF', $g->getLineColorProperty('value'));
    },

    'wave54: gridline style/effect setters are accepted no-ops' => function (): void {
        $g = new GridLines();
        // Chained exactly as a report would; none of these may throw.
        T::ok($g->setLineStyleProperties(2.5, null, 'dash') === $g, 'line style fluent');
        T::ok($g->setGlowProperties(3.0, 'FF0000') === $g, 'glow fluent');
        T::ok($g->setShadowProperties(1) === $g, 'shadow fluent');
        T::ok($g->setSoftEdges(2.0) === $g, 'soft edges fluent');
        T::same(true, $g->getObjectState());
    },

    // ----------------------------------------------------------------- Layout

    'wave54: Layout::setShowVal drives the spec flag' => function (): void {
        $plotArea = new PlotArea(new Layout(), [w54series()]);
        $chart = new Chart('c', null, null, $plotArea);
        T::ok(!\array_key_exists('showValues', w54spec($chart)), 'off by default');

        $chart->getPlotArea()->getLayout()->setShowVal(true);
        T::same(true, w54spec($chart)['showValues']);
    },

    'wave54: getLayout materialises a layout when none was passed' => function (): void {
        $plotArea = new PlotArea(null, [w54series()]);
        $layout = $plotArea->getLayout();
        T::ok($layout instanceof Layout, 'created on demand');
        T::ok($plotArea->getLayout() === $layout, 'and memoised');
    },

    'wave54: Layout data-label toggles round-trip' => function (): void {
        $l = new Layout();
        foreach (['ShowVal', 'ShowCatName', 'ShowSerName', 'ShowPercent', 'ShowLeaderLines', 'ShowLegendKey', 'ShowBubbleSize'] as $name) {
            $set = 'set' . $name;
            $get = 'get' . $name;
            T::same(false, $l->{$get}(), "$name defaults off");
            T::ok($l->{$set}(true) === $l, "$name is fluent");
            T::same(true, $l->{$get}(), "$name round-trips");
        }
    },

    'wave54: Layout geometry round-trips but is documented as unrendered' => function (): void {
        $l = new Layout(['x' => 0.1, 'y' => 0.2, 'w' => 0.5, 'h' => 0.4, 'layoutTarget' => 'inner']);
        T::same(0.1, $l->getXPosition());
        T::same(0.2, $l->getYPosition());
        T::same(0.5, $l->getWidth());
        T::same(0.4, $l->getHeight());
        T::same('inner', $l->getLayoutTarget());
    },

    // ------------------------------------------------------------- ChartColor

    'wave54: ChartColor normalises hex input' => function (): void {
        T::same('FFFFFF', ChartColor::normalise('#ffffff'));
        T::same('FF0000', ChartColor::normalise('ff0000'));
        T::same('00FF00', ChartColor::normalise('00FF00'));
    },

    'wave54: ChartColor stores value, type and alpha' => function (): void {
        $c = new ChartColor();
        $c->setColorProperties('#abcdef', 50, ChartColor::EXCEL_COLOR_TYPE_ARGB);
        T::same('ABCDEF', $c->getValue());
        T::same(ChartColor::EXCEL_COLOR_TYPE_ARGB, $c->getType());
        T::same(50, $c->getAlpha());
    },

    // ------------------------------------------------------ Chart integration

    'wave54: the full constructor shape from the audited app works' => function (): void {
        // Verbatim shape from ProgramVarianceReportManager.
        $series = w54series();
        $series->setPlotDirection(DataSeries::DIRECTION_BAR);
        $plotArea = new PlotArea(new Layout(), [$series]);

        $majorGridLines = new GridLines();
        $majorGridLines->setLineColorProperties('ffffff', 100, GridLines::EXCEL_COLOR_TYPE_ARGB);

        $xAxis = new Axis();
        $xAxis->setAxisOptionsProperties('none');
        $xAxis->setFillParameters('FFFFFF');

        $chart = new Chart(
            'fund_status',
            null,
            null,
            $plotArea,
            true,
            DataSeries::EMPTY_AS_GAP,
            null,
            null,
            $xAxis,
            null,
            $majorGridLines,
        );
        $chart->getPlotArea()->getLayout()->setShowVal(true);
        $chart->setTopLeftPosition('A6');
        $chart->setBottomRightPosition('D20', 0, 0);

        $spec = w54spec($chart);
        T::same('bar', $spec['type'], 'DIRECTION_BAR maps to a bar chart');
        T::same('none', $spec['xAxis']['labels'], 'x axis suppressed');
        T::same(true, $spec['yAxis']['majorGridlines'], 'gridlines land on the Y axis');
        T::same(true, $spec['showValues'], 'layout showVal carried');
    },

    'wave54: gridlines always attach to the Y axis, whichever axis was passed' => function (): void {
        $chart = new Chart(
            'c', null, null, new PlotArea(new Layout(), [w54series()]),
            true, DataSeries::EMPTY_AS_GAP, null, null,
            new Axis(), null, new GridLines(),
        );
        T::same(true, $chart->getChartAxisY()->buildSpec()['majorGridlines'], 'on Y');
        T::same([], $chart->getChartAxisX()->buildSpec(), 'not on X');
    },

    'wave54: bottom-right anchor derives an approximate size' => function (): void {
        $chart = new Chart('c', null, null, new PlotArea(new Layout(), [w54series()]));
        $chart->setTopLeftPosition('A6');
        $chart->setBottomRightPosition('D20');

        $spec = w54spec($chart);
        // A1->D20 spans 3 columns and 14 rows on the OOXML default grid.
        T::same(3 * 64, $spec['width']);
        T::same(14 * 20, $spec['height']);
    },

    'wave54: no bottom-right anchor leaves the size to the engine' => function (): void {
        $chart = new Chart('c', null, null, new PlotArea(new Layout(), [w54series()]));
        $chart->setTopLeftPosition('A1');
        $spec = w54spec($chart);
        T::ok(!\array_key_exists('width', $spec), 'no width');
        T::ok(!\array_key_exists('height', $spec), 'no height');
    },

    'wave54: render() returns false rather than fataling' => function (): void {
        $chart = new Chart('c', null, null, new PlotArea(new Layout(), [w54series()]));
        T::same(false, $chart->render(), 'no PHP-side rasterizer; callers take the no-image branch');
        T::same(false, $chart->render('/tmp/whatever.png'));
    },

    'wave54: getTitle and getTopLeftPosition serve the HTML writer path' => function (): void {
        $title = new Title('Fund status');
        $chart = new Chart('c', $title, null, new PlotArea(new Layout(), [w54series()]));
        $chart->setTopLeftPosition('B4');
        T::ok($chart->getTitle() === $title, 'title round-trips');
        T::same('B4', $chart->getTopLeftPosition()['cell']);
    },

    'wave54: a chart with no axis config still builds the pre-5.4 spec' => function (): void {
        $chart = new Chart('c', new Title('T'), null, new PlotArea(new Layout(), [w54series()]));
        $spec = w54spec($chart);
        T::same('T', $spec['title']);
        T::same(1, \count($spec['series']));
        T::ok(!\array_key_exists('xAxis', $spec), 'no axis block emitted');
        T::ok(!\array_key_exists('yAxis', $spec), 'no axis block emitted');
    },

    'wave54: addChart sends the axis-bearing spec to the extension' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();

        $axis = new Axis();
        $axis->setAxisOptionsProperties('none');
        $chart = new Chart(
            'c', null, null, new PlotArea(new Layout(), [w54series()]),
            true, DataSeries::EMPTY_AS_GAP, null, null, $axis, null, new GridLines(),
        );
        $chart->setTopLeftPosition('A1');
        $ws->addChart($chart);

        $calls = EasyExcelFake::calls('add_chart');
        T::same(1, \count($calls));
        // The fake decodes the JSON the shim encoded, so reaching these keys
        // proves the axis block survived the encode/decode round trip.
        $sent = $calls[0][1][3];
        T::same('none', $sent['xAxis']['labels'], 'axis survives JSON encoding');
        T::same(true, $sent['yAxis']['majorGridlines']);
    },

    // --------------------------------------------------------------- aliasing

    'wave54: the new chart classes resolve through the PhpOffice alias' => function (): void {
        foreach ([
            'PhpOffice\PhpSpreadsheet\Chart\Axis',
            'PhpOffice\PhpSpreadsheet\Chart\GridLines',
            'PhpOffice\PhpSpreadsheet\Chart\Layout',
            'PhpOffice\PhpSpreadsheet\Chart\ChartColor',
        ] as $class) {
            T::ok(\class_exists($class), "$class resolves");
        }
        T::same('none', \constant('PhpOffice\PhpSpreadsheet\Chart\Axis::AXIS_LABELS_NONE'));
    },
];
