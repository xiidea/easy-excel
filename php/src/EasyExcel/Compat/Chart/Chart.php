<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

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

    public function __construct(
        private string $name,
        private ?Title $title = null,
        private ?Legend $legend = null,
        private ?PlotArea $plotArea = null,
        private bool $plotVisibleOnly = true,
        private string $displayBlanksAs = 'gap',
        private ?Title $xAxisLabel = null,
        private ?Title $yAxisLabel = null,
    ) {
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

        return $spec;
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
