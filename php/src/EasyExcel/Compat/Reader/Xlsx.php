<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Reader;

use EasyExcel\Compat\Shared\StreamPath;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Native;

class Xlsx implements IReader
{
    private bool $readDataOnly = false;

    private string $password = '';

    /** easy-excel extra: opens agile-encrypted workbooks. */
    public function setPassword(string $password): static
    {
        $this->password = $password;

        return $this;
    }

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

    private ?IReadFilter $readFilter = null;

    /** Filtered-out cells come back as null from the read APIs (COMPAT.md). */
    public function setReadFilter(IReadFilter $readFilter): static
    {
        $this->readFilter = $readFilter;

        return $this;
    }

    public function getReadFilter(): ?IReadFilter
    {
        return $this->readFilter;
    }

    public function load(string $filename, int $flags = 0): Spreadsheet
    {
        if (StreamPath::isWrapped($filename)) {
            return $this->loadFromStream($filename);
        }
        if (!\is_file($filename)) {
            throw new Exception("File \"$filename\" does not exist.");
        }

        return $this->fromLocalPath($filename);
    }

    /**
     * The extension only opens real filesystem paths, so stream-wrapper URLs
     * (gaufrette://, s3://, ...) are staged into a local temp file first. The
     * temp file is unlinked right away: the native open reads the workbook
     * fully before returning.
     */
    private function loadFromStream(string $url): Spreadsheet
    {
        $in = @\fopen($url, 'rb');
        if ($in === false) {
            throw new Exception("Could not open \"$url\" for reading.");
        }
        $tmp = \tempnam(\sys_get_temp_dir(), 'eexcel');
        if ($tmp === false) {
            \fclose($in);
            throw new Exception('Could not create temporary file');
        }
        $out = \fopen($tmp, 'wb');
        \stream_copy_to_stream($in, $out);
        \fclose($out);
        \fclose($in);

        try {
            return $this->fromLocalPath($tmp);
        } finally {
            @\unlink($tmp);
        }
    }

    private function fromLocalPath(string $path): Spreadsheet
    {
        $spreadsheet = Spreadsheet::fromHandle(Native::open($path, $this->password));
        $spreadsheet->setReadFilter($this->readFilter);

        return $spreadsheet;
    }
}
