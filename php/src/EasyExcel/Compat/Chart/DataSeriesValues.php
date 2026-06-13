<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/** One data source for a chart series (label, category or value range). */
class DataSeriesValues
{
    public const DATASERIES_TYPE_STRING = 'String';
    public const DATASERIES_TYPE_NUMBER = 'Number';

    public function __construct(
        private string $dataType = self::DATASERIES_TYPE_NUMBER,
        private ?string $dataSource = null,
        private ?string $formatCode = null,
        private int $pointCount = 0,
        private array $dataValues = [],
    ) {
    }

    public function getDataType(): string
    {
        return $this->dataType;
    }

    public function getDataSource(): ?string
    {
        return $this->dataSource;
    }

    public function setDataSource(?string $dataSource): static
    {
        $this->dataSource = $dataSource;

        return $this;
    }

    public function getPointCount(): int
    {
        return $this->pointCount;
    }

    /** @return array<int, mixed> */
    public function getDataValues(): array
    {
        return $this->dataValues;
    }
}
