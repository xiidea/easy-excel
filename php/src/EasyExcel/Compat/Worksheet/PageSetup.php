<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Native;

/**
 * Print layout (Phase-2 subset: orientation, paper size, fit-to-page).
 * Each setter pushes the accumulated state; application happens at save.
 */
class PageSetup
{
    public const ORIENTATION_DEFAULT = 'default';
    public const ORIENTATION_LANDSCAPE = 'landscape';
    public const ORIENTATION_PORTRAIT = 'portrait';

    public const PAPERSIZE_LETTER = 1;
    public const PAPERSIZE_TABLOID = 3;
    public const PAPERSIZE_LEGAL = 5;
    public const PAPERSIZE_A3 = 8;
    public const PAPERSIZE_A4 = 9;
    public const PAPERSIZE_A5 = 11;

    private string $orientation = self::ORIENTATION_DEFAULT;
    private int $paperSize = -1;
    private int $fitToWidth = -1;
    private int $fitToHeight = -1;

    public function __construct(private Worksheet $worksheet)
    {
    }

    public function setOrientation(string $orientation): static
    {
        $this->orientation = $orientation;

        return $this->push();
    }

    public function getOrientation(): string
    {
        return $this->orientation;
    }

    public function setPaperSize(int $paperSize): static
    {
        $this->paperSize = $paperSize;

        return $this->push();
    }

    public function getPaperSize(): int
    {
        return $this->paperSize;
    }

    public function setFitToWidth(?int $pages = 1, bool $update = true): static
    {
        $this->fitToWidth = $pages ?? 1;

        return $this->push();
    }

    public function getFitToWidth(): int
    {
        return $this->fitToWidth;
    }

    public function setFitToHeight(?int $pages = 1, bool $update = true): static
    {
        $this->fitToHeight = $pages ?? 1;

        return $this->push();
    }

    public function getFitToHeight(): int
    {
        return $this->fitToHeight;
    }

    public function setFitToPage(bool $fit = true): static
    {
        $this->fitToWidth = $fit ? 1 : -1;
        $this->fitToHeight = $fit ? 1 : -1;

        return $this->push();
    }

    private function push(): static
    {
        Native::pageSetup(
            $this->worksheet->getParent()->getHandle(),
            $this->worksheet->getTitle(),
            $this->orientation === self::ORIENTATION_DEFAULT ? '' : $this->orientation,
            $this->paperSize,
            $this->fitToWidth,
            $this->fitToHeight,
        );

        return $this;
    }
}
