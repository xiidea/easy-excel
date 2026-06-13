<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Native;

/**
 * Print headers/footers; PhpSpreadsheet's placeholder codes (&L/&C/&R, &P,
 * &N, &D, &T, &F, &A, &"font,style") pass through unchanged.
 */
class HeaderFooter
{
    private array $spec = [
        'oddHeader' => '', 'oddFooter' => '',
        'evenHeader' => '', 'evenFooter' => '',
        'firstHeader' => '', 'firstFooter' => '',
        'differentFirst' => false, 'differentOddEven' => false,
    ];

    public function __construct(private Worksheet $worksheet)
    {
    }

    public function setOddHeader(string $value): static
    {
        return $this->set('oddHeader', $value);
    }

    public function getOddHeader(): string
    {
        return $this->spec['oddHeader'];
    }

    public function setOddFooter(string $value): static
    {
        return $this->set('oddFooter', $value);
    }

    public function getOddFooter(): string
    {
        return $this->spec['oddFooter'];
    }

    public function setEvenHeader(string $value): static
    {
        return $this->set('evenHeader', $value);
    }

    public function setEvenFooter(string $value): static
    {
        return $this->set('evenFooter', $value);
    }

    public function setFirstHeader(string $value): static
    {
        return $this->set('firstHeader', $value);
    }

    public function setFirstFooter(string $value): static
    {
        return $this->set('firstFooter', $value);
    }

    public function setDifferentFirst(bool $value): static
    {
        return $this->set('differentFirst', $value);
    }

    public function setDifferentOddEven(bool $value): static
    {
        return $this->set('differentOddEven', $value);
    }

    private function set(string $key, string|bool $value): static
    {
        $this->spec[$key] = $value;
        Native::headerFooter(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->spec,
        );

        return $this;
    }
}
