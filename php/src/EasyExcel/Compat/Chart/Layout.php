<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Chart;

/**
 * Chart layout and data-label options, PhpSpreadsheet style (wave 5.4).
 *
 * Only the data-label toggles map onto excelize: `setShowVal()` becomes the
 * plot area's ShowVal flag, and the sibling toggles map to their ShowCatName /
 * ShowSerName / ShowPercent / ShowLeaderLines counterparts. Manual plot-area
 * geometry (x/y/w/h, layout target, mode) has no excelize equivalent — it is
 * stored and round-trips through the getters, but does not affect the rendered
 * chart (COMPAT.md divergence 30).
 */
class Layout
{
    private ?float $xPos = null;

    private ?float $yPos = null;

    private ?float $width = null;

    private ?float $height = null;

    private string $layoutTarget = '';

    private bool $showLegendKey = false;

    private bool $showVal = false;

    private bool $showCatName = false;

    private bool $showSerName = false;

    private bool $showPercent = false;

    private bool $showBubbleSize = false;

    private bool $showLeaderLines = false;

    /** @param array<string, mixed> $layout */
    public function __construct(array $layout = [])
    {
        foreach ($layout as $key => $value) {
            match ($key) {
                'layoutTarget' => $this->layoutTarget = (string) $value,
                'x' => $this->xPos = (float) $value,
                'y' => $this->yPos = (float) $value,
                'w' => $this->width = (float) $value,
                'h' => $this->height = (float) $value,
                default => null,
            };
        }
    }

    public function setShowVal(bool $value): static
    {
        $this->showVal = $value;

        return $this;
    }

    public function getShowVal(): bool
    {
        return $this->showVal;
    }

    public function setShowLegendKey(bool $value): static
    {
        $this->showLegendKey = $value;

        return $this;
    }

    public function getShowLegendKey(): bool
    {
        return $this->showLegendKey;
    }

    public function setShowCatName(bool $value): static
    {
        $this->showCatName = $value;

        return $this;
    }

    public function getShowCatName(): bool
    {
        return $this->showCatName;
    }

    public function setShowSerName(bool $value): static
    {
        $this->showSerName = $value;

        return $this;
    }

    public function getShowSerName(): bool
    {
        return $this->showSerName;
    }

    public function setShowPercent(bool $value): static
    {
        $this->showPercent = $value;

        return $this;
    }

    public function getShowPercent(): bool
    {
        return $this->showPercent;
    }

    public function setShowBubbleSize(bool $value): static
    {
        $this->showBubbleSize = $value;

        return $this;
    }

    public function getShowBubbleSize(): bool
    {
        return $this->showBubbleSize;
    }

    public function setShowLeaderLines(bool $value): static
    {
        $this->showLeaderLines = $value;

        return $this;
    }

    public function getShowLeaderLines(): bool
    {
        return $this->showLeaderLines;
    }

    /** Stored and round-tripped, but not rendered — see the class docblock. */
    public function setXPosition(float $value): static
    {
        $this->xPos = $value;

        return $this;
    }

    public function getXPosition(): ?float
    {
        return $this->xPos;
    }

    public function setYPosition(float $value): static
    {
        $this->yPos = $value;

        return $this;
    }

    public function getYPosition(): ?float
    {
        return $this->yPos;
    }

    public function setWidth(?float $value): static
    {
        $this->width = $value;

        return $this;
    }

    public function getWidth(): ?float
    {
        return $this->width;
    }

    public function setHeight(?float $value): static
    {
        $this->height = $value;

        return $this;
    }

    public function getHeight(): ?float
    {
        return $this->height;
    }

    public function getLayoutTarget(): string
    {
        return $this->layoutTarget;
    }

    public function setLayoutTarget(string $value): static
    {
        $this->layoutTarget = $value;

        return $this;
    }
}
