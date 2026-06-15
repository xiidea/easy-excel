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
    public const PAPERSIZE_ISO_B4 = 42;
    public const PAPERSIZE_JAPANESE_DOUBLE_POSTCARD = 43;
    public const PAPERSIZE_STANDARD_PAPER_1 = 44;
    public const PAPERSIZE_STANDARD_PAPER_2 = 45;
    public const PAPERSIZE_STANDARD_PAPER_3 = 46;
    public const PAPERSIZE_INVITE_ENVELOPE = 47;
    public const PAPERSIZE_LETTER_EXTRA_PAPER = 48;
    public const PAPERSIZE_LEGAL_EXTRA_PAPER = 49;
    public const PAPERSIZE_TABLOID_EXTRA_PAPER = 50;
    public const PAPERSIZE_A4_EXTRA_PAPER = 51;
    public const PAPERSIZE_LETTER_TRANSVERSE_PAPER = 52;
    public const PAPERSIZE_A4_TRANSVERSE_PAPER = 53;
    public const PAPERSIZE_LETTER_EXTRA_TRANSVERSE_PAPER = 54;
    public const PAPERSIZE_SUPERA_SUPERA_A4_PAPER = 55;
    public const PAPERSIZE_SUPERB_SUPERB_A3_PAPER = 56;
    public const PAPERSIZE_LETTER_PLUS_PAPER = 57;
    public const PAPERSIZE_A4_PLUS_PAPER = 58;
    public const PAPERSIZE_A5_TRANSVERSE_PAPER = 59;
    public const PAPERSIZE_JIS_B5_TRANSVERSE_PAPER = 60;
    public const PAPERSIZE_A3_EXTRA_PAPER = 61;
    public const PAPERSIZE_A5_EXTRA_PAPER = 62;
    public const PAPERSIZE_ISO_B5_EXTRA_PAPER = 63;
    public const PAPERSIZE_A2_PAPER = 64;
    public const PAPERSIZE_A3_TRANSVERSE_PAPER = 65;
    public const PAPERSIZE_A3_EXTRA_TRANSVERSE_PAPER = 66;

    public const SETPRINTRANGE_OVERWRITE = 'O';
    public const SETPRINTRANGE_INSERT = 'I';

    public const PAGEORDER_OVER_THEN_DOWN = 'overThenDown';
    public const PAGEORDER_DOWN_THEN_OVER = 'downThenOver';

    private string $orientation = self::ORIENTATION_DEFAULT;
    private int $paperSize = -1;
    private int $fitToWidth = -1;
    private int $fitToHeight = -1;
    private bool $fitToPage = false;
    private ?int $scale = 100;
    private bool $horizontalCentered = false;
    private bool $verticalCentered = false;
    private ?int $firstPageNumber = null;
    private string $pageOrder = self::PAGEORDER_DOWN_THEN_OVER;

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
        $this->fitToPage = $fit;
        $this->fitToWidth = $fit ? 1 : -1;
        $this->fitToHeight = $fit ? 1 : -1;

        return $this->push();
    }

    public function getFitToPage(): bool
    {
        return $this->fitToPage;
    }

    // --- scale, centering, page order (carried for round-trip; not all wired
    //     into excelize rendering — see COMPAT.md) -------------------------------

    public function getScale(): ?int
    {
        return $this->scale;
    }

    public function setScale(?int $scale, bool $update = true): static
    {
        $this->scale = $scale;
        if ($update) {
            $this->fitToPage = false;
        }

        return $this;
    }

    public function getHorizontalCentered(): bool
    {
        return $this->horizontalCentered;
    }

    public function setHorizontalCentered(bool $value): static
    {
        $this->horizontalCentered = $value;

        return $this;
    }

    public function getVerticalCentered(): bool
    {
        return $this->verticalCentered;
    }

    public function setVerticalCentered(bool $value): static
    {
        $this->verticalCentered = $value;

        return $this;
    }

    public function getFirstPageNumber(): ?int
    {
        return $this->firstPageNumber;
    }

    public function setFirstPageNumber(?int $value): static
    {
        $this->firstPageNumber = $value;

        return $this;
    }

    public function resetFirstPageNumber(): static
    {
        return $this->setFirstPageNumber(null);
    }

    public function getPageOrder(): string
    {
        return $this->pageOrder;
    }

    public function setPageOrder(?string $pageOrder): self
    {
        $this->pageOrder = $pageOrder ?? self::PAGEORDER_DOWN_THEN_OVER;

        return $this;
    }

    // --- defaults (process-wide in PhpSpreadsheet; per-instance is close enough)

    public function getPaperSizeDefault(): int
    {
        return self::PAPERSIZE_LETTER;
    }

    public function setPaperSizeDefault(int $paperSize): void
    {
        // no-op: defaults are not persisted in the streaming model
    }

    public function getOrientationDefault(): string
    {
        return self::ORIENTATION_DEFAULT;
    }

    public function setOrientationDefault(string $orientation): void
    {
        // no-op
    }

    // --- print titles & print area (reserved _xlnm defined names) -------------

    private ?array $repeatRows = null;    // [start, end]
    private ?array $repeatColumns = null; // ["A", "C"]
    private ?string $printArea = null;

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

    public function isRowsToRepeatAtTopSet(): bool
    {
        return $this->repeatRows !== null;
    }

    /** @return array{0: int, 1: int}|array{} */
    public function getRowsToRepeatAtTop(): array
    {
        return $this->repeatRows ?? [];
    }

    public function setColumnsToRepeatAtLeftByStartAndEnd(string $start, string $end): static
    {
        $this->repeatColumns = [$start, $end];

        return $this->pushPrintTitles();
    }

    /** @param array{0: string, 1: string} $columns */
    public function setColumnsToRepeatAtLeft(array $columns): static
    {
        return $this->setColumnsToRepeatAtLeftByStartAndEnd((string) $columns[0], (string) $columns[1]);
    }

    public function isColumnsToRepeatAtLeftSet(): bool
    {
        return $this->repeatColumns !== null;
    }

    /** @return array{0: string, 1: string}|array{} */
    public function getColumnsToRepeatAtLeft(): array
    {
        return $this->repeatColumns ?? [];
    }

    public function setPrintArea(
        string $range,
        int $index = 0,
        string $method = self::SETPRINTRANGE_OVERWRITE,
    ): static {
        $this->printArea = $range;
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

    public function getPrintArea(int $index = 0): string
    {
        return $this->printArea ?? '';
    }

    public function isPrintAreaSet(int $index = 0): bool
    {
        return $this->printArea !== null;
    }

    public function clearPrintArea(int $index = 0): static
    {
        $this->printArea = null;
        Native::definedName(
            $this->worksheet->getParent()->getHandle(),
            '_xlnm.Print_Area',
            '',
            $this->worksheet->getTitle(),
        );

        return $this;
    }

    public function addPrintArea(string $range, int $index = -1): static
    {
        return $this->setPrintArea($range);
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
