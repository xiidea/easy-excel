<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Worksheet;

use EasyExcel\Native;

/**
 * Worksheet protection flags (PhpSpreadsheet polarity: true = action locked).
 * State is pushed on every change; the extension applies it at save.
 */
class Protection
{
    private array $spec = [
        'sheet' => false,
        'password' => '',
        'autoFilter' => false,
        'deleteColumns' => false,
        'deleteRows' => false,
        'formatCells' => false,
        'formatColumns' => false,
        'formatRows' => false,
        'insertColumns' => false,
        'insertHyperlinks' => false,
        'insertRows' => false,
        'objects' => false,
        'pivotTables' => false,
        'scenarios' => false,
        'selectLockedCells' => false,
        'selectUnlockedCells' => false,
        'sort' => false,
    ];

    public function __construct(private Worksheet $worksheet)
    {
    }

    public function setSheet(bool $sheet): static
    {
        return $this->set('sheet', $sheet);
    }

    public function getSheet(): bool
    {
        return $this->spec['sheet'];
    }

    public function setPassword(string $password, bool $alreadyHashed = false): static
    {
        return $this->set('password', $password);
    }

    public function setAutoFilter(bool $locked): static
    {
        return $this->set('autoFilter', $locked);
    }

    public function setDeleteColumns(bool $locked): static
    {
        return $this->set('deleteColumns', $locked);
    }

    public function setDeleteRows(bool $locked): static
    {
        return $this->set('deleteRows', $locked);
    }

    public function setFormatCells(bool $locked): static
    {
        return $this->set('formatCells', $locked);
    }

    public function setFormatColumns(bool $locked): static
    {
        return $this->set('formatColumns', $locked);
    }

    public function setFormatRows(bool $locked): static
    {
        return $this->set('formatRows', $locked);
    }

    public function setInsertColumns(bool $locked): static
    {
        return $this->set('insertColumns', $locked);
    }

    public function setInsertHyperlinks(bool $locked): static
    {
        return $this->set('insertHyperlinks', $locked);
    }

    public function setInsertRows(bool $locked): static
    {
        return $this->set('insertRows', $locked);
    }

    public function setObjects(bool $locked): static
    {
        return $this->set('objects', $locked);
    }

    public function setPivotTables(bool $locked): static
    {
        return $this->set('pivotTables', $locked);
    }

    public function setScenarios(bool $locked): static
    {
        return $this->set('scenarios', $locked);
    }

    public function setSelectLockedCells(bool $locked): static
    {
        return $this->set('selectLockedCells', $locked);
    }

    public function setSelectUnlockedCells(bool $locked): static
    {
        return $this->set('selectUnlockedCells', $locked);
    }

    public function setSort(bool $locked): static
    {
        return $this->set('sort', $locked);
    }

    private function set(string $key, bool|string $value): static
    {
        $this->spec[$key] = $value;
        if ($this->spec['sheet'] === true) {
            Native::protectSheet(
                $this->worksheet->getParent()->getHandle(),
                $this->worksheet->getTitle(),
                $this->spec,
            );
        }

        return $this;
    }
}
