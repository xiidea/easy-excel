<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Native;

/** Sheet view settings (zoom, RTL); gridlines live on the Worksheet. */
class SheetView
{
    private array $spec = [];

    public function __construct(private Worksheet $worksheet)
    {
    }

    public function setZoomScale(?int $zoomScale): static
    {
        return $this->set('zoomScale', $zoomScale === null ? null : (float) $zoomScale);
    }

    public function getZoomScale(): ?int
    {
        return isset($this->spec['zoomScale']) ? (int) $this->spec['zoomScale'] : null;
    }

    public function setZoomScaleNormal(?int $zoomScale): static
    {
        return $this; // screen-only hint, not persisted by excelize
    }

    public function setRightToLeft(bool $rtl): static
    {
        return $this->set('rightToLeft', $rtl);
    }

    /** @internal shared push for Worksheet::setShowGridlines / getTabColor */
    public function pushExtra(string $key, mixed $value): void
    {
        $this->set($key, $value);
    }

    private function set(string $key, mixed $value): static
    {
        if ($value === null) {
            unset($this->spec[$key]);
        } else {
            $this->spec[$key] = $value;
        }
        Native::sheetView(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->spec,
        );

        return $this;
    }
}
