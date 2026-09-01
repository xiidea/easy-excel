<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/**
 * Chart gridlines, PhpSpreadsheet style (wave 5.4).
 *
 * Constructing a GridLines and handing it to Chart turns the corresponding
 * gridlines on — that is the part excelize models (`ChartAxis.MajorGridLines`
 * / `MinorGridLines`). The line styling setters below are accepted for source
 * compatibility but not rendered: excelize exposes no gridline line format
 * (COMPAT.md divergence 30).
 */
class GridLines extends ChartColor
{
    private bool $objectState = true;

    private ?string $lineColor = null;

    public function getObjectState(): bool
    {
        return $this->objectState;
    }

    public function activate(): static
    {
        $this->objectState = true;

        return $this;
    }

    /**
     * Accepted, not rendered — excelize has no gridline line-format model.
     * The colour is retained so getLineColorProperty() round-trips.
     */
    public function setLineColorProperties(?string $value, null|float|int|string $alpha = null, ?string $colorType = null): static
    {
        if ($value !== null) {
            $this->lineColor = ChartColor::normalise($value);
        }
        $this->setColorProperties($value, $alpha, $colorType);

        return $this;
    }

    public function getLineColorProperty(string $propertyName = 'value'): ?string
    {
        return match ($propertyName) {
            'value' => $this->lineColor,
            'type' => $this->getType(),
            default => null,
        };
    }

    /** Accepted no-op: line width/style are not modelled by excelize. */
    public function setLineStyleProperties(
        null|float|int|string $lineWidth = null,
        ?string $compoundType = null,
        ?string $dashType = null,
        ?string $capType = null,
        null|float|int|string $joinType = null,
        ?string $headArrowType = null,
        ?string $headArrowSize = null,
        ?string $endArrowType = null,
        ?string $endArrowSize = null,
        ?string $headArrowWidth = null,
        ?string $headArrowLength = null,
        ?string $endArrowWidth = null,
        ?string $endArrowLength = null,
    ): static {
        return $this;
    }

    /** Accepted no-op: glow/shadow/soft-edge effects are not modelled. */
    public function setGlowProperties(float $size, ?string $colorValue = null, null|float|int|string $colorAlpha = null, ?string $colorType = null): static
    {
        return $this;
    }

    public function setShadowProperties(
        ?int $presets = null,
        ?string $colorValue = null,
        ?string $colorType = null,
        null|float|int|string $colorAlpha = null,
        ?float $blur = null,
        ?int $angle = null,
        ?float $distance = null,
    ): static {
        return $this;
    }

    public function setSoftEdges(?float $size): static
    {
        return $this;
    }
}
