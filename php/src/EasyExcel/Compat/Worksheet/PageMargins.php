<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Native;

/** Print margins in inches (PhpSpreadsheet units). */
class PageMargins
{
    private array $spec = [
        'top' => -1.0, 'bottom' => -1.0, 'left' => -1.0,
        'right' => -1.0, 'header' => -1.0, 'footer' => -1.0,
    ];

    public function __construct(private Worksheet $worksheet)
    {
    }

    public function setTop(float $value): static
    {
        return $this->set('top', $value);
    }

    public function getTop(): float
    {
        return $this->spec['top'];
    }

    public function setBottom(float $value): static
    {
        return $this->set('bottom', $value);
    }

    public function getBottom(): float
    {
        return $this->spec['bottom'];
    }

    public function setLeft(float $value): static
    {
        return $this->set('left', $value);
    }

    public function getLeft(): float
    {
        return $this->spec['left'];
    }

    public function setRight(float $value): static
    {
        return $this->set('right', $value);
    }

    public function getRight(): float
    {
        return $this->spec['right'];
    }

    public function setHeader(float $value): static
    {
        return $this->set('header', $value);
    }

    public function setFooter(float $value): static
    {
        return $this->set('footer', $value);
    }

    private function set(string $key, float $value): static
    {
        $this->spec[$key] = $value;
        Native::pageMargins(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->spec,
        );

        return $this;
    }
}
