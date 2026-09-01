<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/** Holds the chart's data series and its layout. */
class PlotArea
{
    /** @param list<DataSeries> $plotSeries */
    public function __construct(private ?Layout $layout = null, private array $plotSeries = [])
    {
    }

    /**
     * The layout, created on demand so `getPlotArea()->getLayout()->setShowVal()`
     * works even when the chart was built with `new PlotArea(null, …)`.
     */
    public function getLayout(): Layout
    {
        return $this->layout ??= new Layout();
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
