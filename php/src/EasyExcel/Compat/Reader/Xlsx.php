<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Reader;

use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Native;

class Xlsx
{
    private bool $readDataOnly = false;

    public function setReadDataOnly(bool $readDataOnly): static
    {
        // values-only iteration is already the extension's fast path;
        // the flag is accepted for API parity
        $this->readDataOnly = $readDataOnly;

        return $this;
    }

    public function getReadDataOnly(): bool
    {
        return $this->readDataOnly;
    }

    public function canRead(string $filename): bool
    {
        return \is_readable($filename)
            && \in_array(\strtolower(\pathinfo($filename, PATHINFO_EXTENSION)), ['xlsx', 'xlsm', 'xltx', 'xltm'], true);
    }

    public function load(string $filename, int $flags = 0): Spreadsheet
    {
        if (!\is_file($filename)) {
            throw new Exception("File \"$filename\" does not exist.");
        }

        return Spreadsheet::fromHandle(Native::open($filename));
    }
}
