<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;

use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Native;

class Csv
{
    private string $delimiter = ',';
    private string $enclosure = '"';
    private string $lineEnding = PHP_EOL;
    private bool $useBOM = false;
    private int $sheetIndex = 0;
    private bool $sanitizeFormulas = false; // easy-excel extra: OWASP injection guard

    public function __construct(private Spreadsheet $spreadsheet)
    {
    }

    public function setDelimiter(string $delimiter): static
    {
        $this->delimiter = $delimiter;

        return $this;
    }

    public function getDelimiter(): string
    {
        return $this->delimiter;
    }

    public function setEnclosure(string $enclosure = '"'): static
    {
        if ($enclosure !== '"') {
            throw new Exception('easy-excel: only \'"\' enclosure is supported (COMPAT.md)');
        }
        $this->enclosure = $enclosure;

        return $this;
    }

    public function getEnclosure(): string
    {
        return $this->enclosure;
    }

    public function setLineEnding(string $lineEnding): static
    {
        $this->lineEnding = $lineEnding;

        return $this;
    }

    public function getLineEnding(): string
    {
        return $this->lineEnding;
    }

    public function setUseBOM(bool $useBOM): static
    {
        $this->useBOM = $useBOM;

        return $this;
    }

    public function getUseBOM(): bool
    {
        return $this->useBOM;
    }

    public function setSheetIndex(int $sheetIndex): static
    {
        $this->sheetIndex = $sheetIndex;

        return $this;
    }

    public function getSheetIndex(): int
    {
        return $this->sheetIndex;
    }

    public function setSanitizeFormulas(bool $sanitize): static
    {
        $this->sanitizeFormulas = $sanitize;

        return $this;
    }

    public function save(string $filename, int $flags = 0): void
    {
        $this->spreadsheet->flushAll();
        $handle = $this->spreadsheet->getHandle();
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex)->getTitle();
        $crlf = $this->lineEnding === "\r\n";

        if (!\str_starts_with($filename, 'php://')) {
            Native::saveCsv($handle, $filename, $sheet, $this->delimiter, $crlf, $this->useBOM, $this->sanitizeFormulas);

            return;
        }

        $tmp = \tempnam(\sys_get_temp_dir(), 'eexcel');
        if ($tmp === false) {
            throw new Exception('Could not create temporary file');
        }
        try {
            Native::saveCsv($handle, $tmp, $sheet, $this->delimiter, $crlf, $this->useBOM, $this->sanitizeFormulas);
            $out = \fopen($filename, 'wb');
            if ($out === false) {
                throw new Exception("Could not open $filename for writing");
            }
            $in = \fopen($tmp, 'rb');
            \stream_copy_to_stream($in, $out);
            \fclose($in);
            \fclose($out);
        } finally {
            @\unlink($tmp);
        }
    }
}
