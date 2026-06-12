<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Cell;

use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/**
 * Cell hyperlink. When obtained via Cell::getHyperlink() it is bound to the
 * cell and setters apply immediately; standalone instances are applied via
 * Worksheet::setHyperlink().
 */
class Hyperlink
{
    private ?Worksheet $worksheet = null;
    private string $cell = '';

    public function __construct(private string $url = '', private string $tooltip = '')
    {
    }

    /** @internal */
    public function bind(Worksheet $worksheet, string $cell): static
    {
        $this->worksheet = $worksheet;
        $this->cell = $cell;

        return $this;
    }

    public function setUrl(string $url): static
    {
        $this->url = $url;
        $this->push();

        return $this;
    }

    public function getUrl(): string
    {
        return $this->url;
    }

    public function setTooltip(string $tooltip): static
    {
        $this->tooltip = $tooltip;
        if ($this->url !== '') {
            $this->push();
        }

        return $this;
    }

    public function getTooltip(): string
    {
        return $this->tooltip;
    }

    public function isInternal(): bool
    {
        return \str_starts_with($this->url, 'sheet://');
    }

    private function push(): void
    {
        if ($this->worksheet === null) {
            return; // standalone instance, applied by Worksheet::setHyperlink()
        }
        Native::setHyperlink(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->cell,
            $this->url,
            $this->tooltip,
        );
    }
}
