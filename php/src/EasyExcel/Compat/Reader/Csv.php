<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Reader;

use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;

class Csv
{
    private string $delimiter = ',';
    private string $enclosure = '"';
    private int $sheetIndex = 0;

    public function setDelimiter(string $delimiter): static
    {
        $this->delimiter = $delimiter;

        return $this;
    }

    public function getDelimiter(): string
    {
        return $this->delimiter;
    }

    public function setEnclosure(string $enclosure): static
    {
        $this->enclosure = $enclosure;

        return $this;
    }

    public function getEnclosure(): string
    {
        return $this->enclosure;
    }

    public function setSheetIndex(int $sheetIndex): static
    {
        $this->sheetIndex = $sheetIndex;

        return $this;
    }

    public function canRead(string $filename): bool
    {
        return \is_readable($filename);
    }

    /**
     * Streams the CSV into the native workbook in 1k-row chunks: constant
     * PHP memory, sequential rows keep the Go side in streaming mode.
     */
    public function load(string $filename, int $flags = 0): Spreadsheet
    {
        $fh = @\fopen($filename, 'rb');
        if ($fh === false) {
            throw new Exception("Could not open file $filename for reading.");
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        try {
            // skip a UTF-8 BOM if present
            $bom = \fread($fh, 3);
            if ($bom !== "\xEF\xBB\xBF") {
                \rewind($fh);
            }
            $chunk = [];
            $row = 1;
            while (($data = \fgetcsv($fh, 0, $this->delimiter, $this->enclosure, '')) !== false) {
                $chunk[] = $data;
                if (\count($chunk) >= 1000) {
                    $sheet->fromArray($chunk, null, 'A' . $row, true);
                    $row += \count($chunk);
                    $chunk = [];
                }
            }
            if ($chunk !== []) {
                $sheet->fromArray($chunk, null, 'A' . $row, true);
            }
        } finally {
            \fclose($fh);
        }

        return $spreadsheet;
    }
}
