<?php

declare(strict_types=1);

namespace EasyExcel\Compat;

use EasyExcel\Compat\Worksheet\Worksheet;
use EasyExcel\Native;

/**
 * PhpSpreadsheet-compatible workbook holding only an int64 extension handle;
 * all data lives in Go. Destruction (or disconnectWorksheets(), the
 * PhpSpreadsheet idiom) releases the native workbook; the extension's idle
 * TTL is the backstop (PLAN.md §7.6).
 */
class Spreadsheet
{
    private int $handle;

    /** @var list<Worksheet> in workbook order */
    private array $worksheets = [];

    private bool $connected = true;

    public function __construct()
    {
        $this->handle = Native::newWorkbook();
        $this->worksheets[] = new Worksheet($this, 'Worksheet');
    }

    /** @internal wrap an already-open native workbook (IOFactory::load) */
    public static function fromHandle(int $handle): static
    {
        /** @var static $s */
        $s = (new \ReflectionClass(static::class))->newInstanceWithoutConstructor();
        $s->handle = $handle;
        $s->connected = true;
        foreach (Native::sheets($handle) as $name) {
            $s->worksheets[] = new Worksheet($s, $name);
        }

        return $s;
    }

    /** @internal */
    public function getHandle(): int
    {
        if (!$this->connected) {
            throw new Exception('Spreadsheet has been disconnected');
        }

        return $this->handle;
    }

    public function getActiveSheet(): Worksheet
    {
        [$pos] = Native::activeSheet($this->getHandle());

        return $this->worksheets[$pos] ?? $this->worksheets[0];
    }

    public function getActiveSheetIndex(): int
    {
        [$pos] = Native::activeSheet($this->getHandle());

        return $pos;
    }

    public function setActiveSheetIndex(int $index): Worksheet
    {
        Native::setActiveSheet($this->getHandle(), $index);

        return $this->worksheets[$index];
    }

    public function setActiveSheetIndexByName(string $worksheetName): Worksheet
    {
        foreach ($this->worksheets as $i => $ws) {
            if ($ws->getTitle() === $worksheetName) {
                return $this->setActiveSheetIndex($i);
            }
        }

        throw new Exception("Workbook does not contain sheet: $worksheetName");
    }

    public function createSheet(?int $sheetIndex = null): Worksheet
    {
        if ($sheetIndex !== null && $sheetIndex !== \count($this->worksheets)) {
            throw new Exception('easy-excel: inserting sheets at arbitrary positions is not supported yet (COMPAT.md)');
        }
        $name = $this->nextSheetName();
        Native::addSheet($this->getHandle(), $name);
        $ws = new Worksheet($this, $name);
        $this->worksheets[] = $ws;

        return $ws;
    }

    public function getSheet(int $index): Worksheet
    {
        return $this->worksheets[$index]
            ?? throw new Exception("Your requested sheet index: $index is out of bounds.");
    }

    public function getSheetByName(string $name): ?Worksheet
    {
        foreach ($this->worksheets as $ws) {
            if ($ws->getTitle() === $name) {
                return $ws;
            }
        }

        return null;
    }

    public function getSheetCount(): int
    {
        return \count($this->worksheets);
    }

    /** @return list<string> */
    public function getSheetNames(): array
    {
        return \array_map(static fn (Worksheet $ws): string => $ws->getTitle(), $this->worksheets);
    }

    /** @return list<Worksheet> */
    public function getAllSheets(): array
    {
        return $this->worksheets;
    }

    public function getIndex(Worksheet $worksheet): int
    {
        foreach ($this->worksheets as $i => $ws) {
            if ($ws === $worksheet) {
                return $i;
            }
        }

        throw new Exception('Sheet does not exist.');
    }

    public function removeSheetByIndex(int $sheetIndex): void
    {
        $ws = $this->getSheet($sheetIndex);
        $ws->flush();
        Native::deleteSheet($this->getHandle(), $ws->getTitle());
        \array_splice($this->worksheets, $sheetIndex, 1);
    }

    /** @internal flush every sheet's write-behind buffer */
    public function flushAll(): void
    {
        foreach ($this->worksheets as $ws) {
            $ws->flush();
        }
    }

    /** Releases the native workbook, like PhpSpreadsheet's memory-free idiom. */
    public function disconnectWorksheets(): void
    {
        if ($this->connected) {
            $this->connected = false;
            Native::close($this->handle);
        }
    }

    public function garbageCollect(): static
    {
        return $this;
    }

    public function __destruct()
    {
        $this->disconnectWorksheets();
    }

    private function nextSheetName(): string
    {
        $names = $this->getSheetNames();
        $i = \count($names);
        do {
            $name = 'Worksheet' . $i;
            ++$i;
        } while (\in_array($name, $names, true));

        return $name;
    }
}
