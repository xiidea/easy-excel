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
    public const PAPERSIZE_LETTER_SMALL = 2;
    public const PAPERSIZE_TABLOID = 3;
    public const PAPERSIZE_LEDGER = 4;
    public const PAPERSIZE_LEGAL = 5;
    public const PAPERSIZE_STATEMENT = 6;
    public const PAPERSIZE_EXECUTIVE = 7;
    public const PAPERSIZE_A3 = 8;
    public const PAPERSIZE_A4 = 9;
    public const PAPERSIZE_A4_SMALL = 10;
    public const PAPERSIZE_A5 = 11;
    public const PAPERSIZE_B4 = 12;
    public const PAPERSIZE_B5 = 13;
    public const PAPERSIZE_FOLIO = 14;
    public const PAPERSIZE_QUARTO = 15;
    public const PAPERSIZE_STANDARD_1 = 16;
    public const PAPERSIZE_STANDARD_2 = 17;
    public const PAPERSIZE_NOTE = 18;
    public const PAPERSIZE_NO9_ENVELOPE = 19;
    public const PAPERSIZE_NO10_ENVELOPE = 20;
    public const PAPERSIZE_NO11_ENVELOPE = 21;
    public const PAPERSIZE_NO12_ENVELOPE = 22;
    public const PAPERSIZE_NO14_ENVELOPE = 23;
    public const PAPERSIZE_C = 24;
    public const PAPERSIZE_D = 25;
    public const PAPERSIZE_E = 26;
    public const PAPERSIZE_DL_ENVELOPE = 27;
    public const PAPERSIZE_C5_ENVELOPE = 28;
    public const PAPERSIZE_C3_ENVELOPE = 29;
    public const PAPERSIZE_C4_ENVELOPE = 30;
    public const PAPERSIZE_C6_ENVELOPE = 31;
    public const PAPERSIZE_C65_ENVELOPE = 32;
    public const PAPERSIZE_B4_ENVELOPE = 33;
    public const PAPERSIZE_B5_ENVELOPE = 34;
    public const PAPERSIZE_B6_ENVELOPE = 35;
    public const PAPERSIZE_ITALY_ENVELOPE = 36;
    public const PAPERSIZE_MONARCH_ENVELOPE = 37;
    public const PAPERSIZE_6_3_4_ENVELOPE = 38;
    public const PAPERSIZE_US_STANDARD_FANFOLD = 39;
    public const PAPERSIZE_GERMAN_STANDARD_FANFOLD = 40;
    public const PAPERSIZE_GERMAN_LEGAL_FANFOLD = 41;

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

    // --- print titles & print area (reserved _xlnm defined names) -------------

    private ?array $repeatRows = null;    // [start, end]
    private ?array $repeatColumns = null; // ["A", "C"]

    public function setRowsToRepeatAtTopByStartAndEnd(int $start, int $end): static
    {
        $this->repeatRows = [$start, $end];

        return $this->pushPrintTitles();
    }

    /** @param array{0: int, 1: int} $rows */
    public function setRowsToRepeatAtTop(array $rows): static
    {
        return $this->setRowsToRepeatAtTopByStartAndEnd((int) $rows[0], (int) $rows[1]);
    }

    public function setColumnsToRepeatAtLeftByStartAndEnd(string $start, string $end): static
    {
        $this->repeatColumns = [$start, $end];

        return $this->pushPrintTitles();
    }

    public function setPrintArea(string $range): static
    {
        // absolute sheet-qualified ref, like PhpSpreadsheet writes it
        $abs = \preg_replace('/([A-Z]+)(\d+)/', '\$$1\$$2', \strtoupper($range));
        Native::definedName(
            $this->worksheet->getParent()->getHandle(),
            '_xlnm.Print_Area',
            $this->quotedTitle() . '!' . $abs,
            $this->worksheet->getTitle(),
        );

        return $this;
    }

    private function pushPrintTitles(): static
    {
        $parts = [];
        if ($this->repeatColumns !== null) {
            $parts[] = \sprintf('%s!$%s:$%s', $this->quotedTitle(), $this->repeatColumns[0], $this->repeatColumns[1]);
        }
        if ($this->repeatRows !== null) {
            $parts[] = \sprintf('%s!$%d:$%d', $this->quotedTitle(), $this->repeatRows[0], $this->repeatRows[1]);
        }
        Native::definedName(
            $this->worksheet->getParent()->getHandle(),
            '_xlnm.Print_Titles',
            \implode(',', $parts),
            $this->worksheet->getTitle(),
        );

        return $this;
    }

    private function quotedTitle(): string
    {
        $title = $this->worksheet->getTitle();
        if (\str_contains($title, ' ') || \str_contains($title, "'")) {
            return "'" . \str_replace("'", "''", $title) . "'";
        }

        return $title;
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
