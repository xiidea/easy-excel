<?php

declare(strict_types=1);

namespace EasyExcel\Compat;

use EasyExcel\Compat\Reader\Csv as CsvReader;
use EasyExcel\Compat\Reader\Xlsx as XlsxReader;
use EasyExcel\Compat\Writer\Csv as CsvWriter;
use EasyExcel\Compat\Writer\Xlsx as XlsxWriter;

abstract class IOFactory
{
    public const READER_XLSX = 'Xlsx';
    public const READER_CSV = 'Csv';
    public const WRITER_XLSX = 'Xlsx';
    public const WRITER_CSV = 'Csv';

    public static function createWriter(Spreadsheet $spreadsheet, string $writerType): XlsxWriter|CsvWriter
    {
        return match ($writerType) {
            self::WRITER_XLSX => new XlsxWriter($spreadsheet),
            self::WRITER_CSV => new CsvWriter($spreadsheet),
            default => throw new Exception(
                "easy-excel: writer \"$writerType\" is not supported yet (COMPAT.md lists supported formats)"
            ),
        };
    }

    public static function createReader(string $readerType): XlsxReader|CsvReader
    {
        return match ($readerType) {
            self::READER_XLSX => new XlsxReader(),
            self::READER_CSV => new CsvReader(),
            default => throw new Exception(
                "easy-excel: reader \"$readerType\" is not supported yet (COMPAT.md lists supported formats)"
            ),
        };
    }

    public static function load(string $filename, int $flags = 0, ?array $readers = null): Spreadsheet
    {
        return self::createReader(self::identify($filename))->load($filename);
    }

    public static function identify(string $filename, ?array $readers = null): string
    {
        $ext = \strtolower(\pathinfo($filename, PATHINFO_EXTENSION));

        return match ($ext) {
            'xlsx', 'xlsm', 'xltx', 'xltm' => self::READER_XLSX,
            'csv', 'tsv' => self::READER_CSV,
            default => throw new Exception("Unable to identify a reader for this file: $filename"),
        };
    }
}
