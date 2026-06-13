<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/** Holds the chart's data series (the layout argument is ignored). */
class PlotArea
{
    /** @param list<DataSeries> $plotSeries */
    public function __construct(private mixed $layout = null, private array $plotSeries = [])
    {
    }

    /** @return list<DataSeries> */
    public function getPlotGroup(): array
    {
        return $this->plotSeries;
    }

    public function getPlotGroupByIndex(int $index): DataSeries
    {
        return $this->plotSeries[$index];
    }
}
