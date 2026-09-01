<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Compat\Exception;

/**
 * PhpSpreadsheet-compatible chart facade (wave 4.4). Maps the
 * Chart/DataSeries/DataSeriesValues object model onto easy-excel's native
 * declarative chart spec (extension/compat/chart.go). Supported plot types:
 * bar/column (clustered + stacked), line, area, pie, doughnut, scatter, radar.
 */
class Chart
{
    private string $topLeftCell = 'A1';

    private string $bottomRightCell = '';

    private Axis $xAxis;

    private Axis $yAxis;

    public function __construct(
        private string $name,
        private ?Title $title = null,
        private ?Legend $legend = null,
        private ?PlotArea $plotArea = null,
        private bool $plotVisibleOnly = true,
        private string $displayBlanksAs = 'gap',
        private ?Title $xAxisLabel = null,
        private ?Title $yAxisLabel = null,
        ?Axis $xAxis = null,
        ?Axis $yAxis = null,
        ?GridLines $majorGridlines = null,
        ?GridLines $minorGridlines = null,
    ) {
        $this->xAxis = $xAxis ?? new Axis();
        $this->yAxis = $yAxis ?? new Axis();
        // PhpSpreadsheet attaches both gridline sets to the *value* (y) axis,
        // regardless of which axis object was passed in.
        if ($majorGridlines !== null) {
            $this->yAxis->setMajorGridlines($majorGridlines);
        }
        if ($minorGridlines !== null) {
            $this->yAxis->setMinorGridlines($minorGridlines);
        }
    }

    public function getChartAxisX(): Axis
    {
        return $this->xAxis;
    }

    public function getChartAxisY(): Axis
    {
        return $this->yAxis;
    }

    public function getPlotArea(): ?PlotArea
    {
        return $this->plotArea;
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function setTopLeftPosition(string $cell): static
    {
        $this->topLeftCell = \explode(':', $cell)[0];

        return $this;
    }

    public function getTopLeftCell(): string
    {
        return $this->topLeftCell;
    }

    public function setTopLeftCell(string $cell): static
    {
        return $this->setTopLeftPosition($cell);
    }

    public function getTopLeftPosition(): array
    {
        return ['cell' => $this->topLeftCell, 'xOffset' => 0, 'yOffset' => 0];
    }

    /**
     * Records the chart's bottom-right anchor.
     *
     * excelize sizes charts by explicit width/height rather than a second
     * anchor, so the cell is stored and the span is converted to an
     * approximate pixel size at spec-build time (COMPAT.md divergence 30).
     */
    public function setBottomRightPosition(string $cell, int $xOffset = 0, int $yOffset = 0): static
    {
        $this->bottomRightCell = \explode(':', $cell)[0];

        return $this;
    }

    public function getBottomRightCell(): string
    {
        return $this->bottomRightCell;
    }

    public function getTitle(): ?Title
    {
        return $this->title;
    }

    public function getLegend(): ?Legend
    {
        return $this->legend;
    }

    /**
     * Not supported: rendering a chart to an image needs a PHP-side renderer
     * (PhpSpreadsheet uses JpGraph). easy-excel emits charts as native Excel
     * chart parts and never rasterizes them, so there is nothing to render.
     *
     * Returns false — the same value PhpSpreadsheet returns when no chart
     * renderer is configured — so callers take their existing "no image"
     * branch instead of fataling.
     */
    public function render(?string $outputDestination = null): bool
    {
        return false;
    }

    /** @internal builds the native chart spec for Native::addChart */
    public function buildSpec(): array
    {
        if ($this->plotArea === null) {
            throw new Exception('easy-excel: chart needs a plot area');
        }
        $groups = $this->plotArea->getPlotGroup();
        if ($groups === []) {
            throw new Exception('easy-excel: chart needs at least one data series group');
        }
        $group = $groups[0];

        $spec = ['type' => $this->nativeType($group), 'series' => []];
        $categories = $group->getPlotCategories();
        $labels = $group->getPlotLabels();
        foreach ($group->getPlotValues() as $i => $values) {
            $category = $categories[$i] ?? $categories[0] ?? null;
            $label = $labels[$i] ?? null;
            $spec['series'][] = [
                'name' => $label?->getDataSource() ?? '',
                'categories' => $category?->getDataSource() ?? '',
                'values' => $values->getDataSource() ?? '',
            ];
        }
        if ($this->title !== null && $this->title->getCaptionText() !== '') {
            $spec['title'] = $this->title->getCaptionText();
        }
        if ($this->legend !== null) {
            $spec['legend'] = ['position' => $this->legend->nativePosition()];
        }
        if ($this->xAxisLabel !== null && $this->xAxisLabel->getCaptionText() !== '') {
            $spec['xAxisTitle'] = $this->xAxisLabel->getCaptionText();
        }
        if ($this->yAxisLabel !== null && $this->yAxisLabel->getCaptionText() !== '') {
            $spec['yAxisTitle'] = $this->yAxisLabel->getCaptionText();
        }
        if (($axis = $this->xAxis->buildSpec()) !== []) {
            $spec['xAxis'] = $axis;
        }
        if (($axis = $this->yAxis->buildSpec()) !== []) {
            $spec['yAxis'] = $axis;
        }
        if ($this->plotArea->getLayout()->getShowVal()) {
            $spec['showValues'] = true;
        }
        if (($size = $this->sizeFromAnchors()) !== null) {
            [$spec['width'], $spec['height']] = $size;
        }

        return $spec;
    }

    /**
     * Derive an approximate pixel size from the top-left/bottom-right anchors.
     *
     * excelize sizes a chart by width/height, not by a second anchor, so the
     * span is converted with the OOXML default grid (64px per column, 20px per
     * row). Approximate by construction — see COMPAT.md divergence 30.
     *
     * @return null|array{int, int}
     */
    private function sizeFromAnchors(): ?array
    {
        if ($this->bottomRightCell === '') {
            return null;
        }
        try {
            [$c1, $r1] = Coordinate::indexesFromString($this->topLeftCell);
            [$c2, $r2] = Coordinate::indexesFromString($this->bottomRightCell);
        } catch (\Throwable) {
            return null;
        }
        $cols = \max(1, $c2 - $c1);
        $rows = \max(1, $r2 - $r1);

        return [$cols * 64, $rows * 20];
    }

    private function nativeType(DataSeries $group): string
    {
        $stacked = \in_array($group->getPlotGrouping(), [
            DataSeries::GROUPING_STACKED,
            DataSeries::GROUPING_PERCENT_STACKED,
        ], true);

        return match ($group->getPlotType()) {
            DataSeries::TYPE_BARCHART, DataSeries::TYPE_BARCHART_3D => match (true) {
                $group->getPlotDirection() === DataSeries::DIRECTION_BAR => $stacked ? 'barStacked' : 'bar',
                default => $stacked ? 'colStacked' : 'col',
            },
            DataSeries::TYPE_LINECHART => 'line',
            DataSeries::TYPE_AREACHART => 'area',
            DataSeries::TYPE_PIECHART => 'pie',
            DataSeries::TYPE_DOUGHNUTCHART => 'doughnut',
            DataSeries::TYPE_SCATTERCHART => 'scatter',
            DataSeries::TYPE_RADARCHART => 'radar',
            default => throw new Exception('easy-excel: unsupported chart plot type ' . $group->getPlotType()),
        };
    }
}
