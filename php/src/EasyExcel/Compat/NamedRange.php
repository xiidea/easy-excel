<?php

declare(strict_types=1);

namespace EasyExcel\Compat;

use EasyExcel\Compat\Worksheet\Worksheet;

/** Named range, registered through Spreadsheet::addNamedRange(). */
class NamedRange
{
    public function __construct(
        private string $name,
        private ?Worksheet $worksheet = null,
        private string $range = 'A1',
        private bool $localOnly = false,
    ) {
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getRange(): string
    {
        return $this->range;
    }

    public function getWorksheet(): ?Worksheet
    {
        return $this->worksheet;
    }

    public function getLocalOnly(): bool
    {
        return $this->localOnly;
    }

    /** @internal absolute sheet-qualified reference for the extension */
    public function getRefersTo(): string
    {
        $range = \ltrim($this->range, '=');
        if (\str_contains($range, '!') || $this->worksheet === null) {
            return $range;
        }
        $title = $this->worksheet->getTitle();
        if (\str_contains($title, ' ') || \str_contains($title, "'")) {
            $title = "'" . \str_replace("'", "''", $title) . "'";
        }

        return $title . '!' . $range;
    }
}
