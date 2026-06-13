<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/** A chart's plot series set (PhpSpreadsheet DataSeries). */
class DataSeries
{
    public const TYPE_BARCHART = 'barChart';
    public const TYPE_BARCHART_3D = 'bar3DChart';
    public const TYPE_LINECHART = 'lineChart';
    public const TYPE_AREACHART = 'areaChart';
    public const TYPE_PIECHART = 'pieChart';
    public const TYPE_DOUGHNUTCHART = 'doughnutChart';
    public const TYPE_SCATTERCHART = 'scatterChart';
    public const TYPE_RADARCHART = 'radarChart';

    public const GROUPING_CLUSTERED = 'clustered';
    public const GROUPING_STACKED = 'stacked';
    public const GROUPING_PERCENT_STACKED = 'percentStacked';
    public const GROUPING_STANDARD = 'standard';

    public const DIRECTION_BAR = 'bar';
    public const DIRECTION_COL = 'col';

    private string $plotDirection = self::DIRECTION_COL;

    /**
     * @param list<DataSeriesValues> $plotLabel
     * @param list<DataSeriesValues> $plotCategory
     * @param list<DataSeriesValues> $plotValues
     */
    public function __construct(
        private ?string $plotType = null,
        private ?string $plotGrouping = null,
        private array $plotOrder = [],
        private array $plotLabel = [],
        private array $plotCategory = [],
        private array $plotValues = [],
    ) {
    }

    public function getPlotType(): ?string
    {
        return $this->plotType;
    }

    public function getPlotGrouping(): ?string
    {
        return $this->plotGrouping;
    }

    public function setPlotDirection(string $direction): static
    {
        $this->plotDirection = $direction;

        return $this;
    }

    public function getPlotDirection(): string
    {
        return $this->plotDirection;
    }

    /** @return list<DataSeriesValues> */
    public function getPlotValues(): array
    {
        return $this->plotValues;
    }

    /** @return list<DataSeriesValues> */
    public function getPlotLabels(): array
    {
        return $this->plotLabel;
    }

    /** @return list<DataSeriesValues> */
    public function getPlotCategories(): array
    {
        return $this->plotCategory;
    }
}
