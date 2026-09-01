<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Reader;

use EasyExcel\Compat\Spreadsheet;

/**
 * Reader contract, PhpSpreadsheet style (wave 5.2). Scoped to the members
 * both Compat readers genuinely share — consumers type-hint against this and
 * IOFactory::createReader() returns it.
 */
interface IReader
{
    /** Read data only, ignoring styles/formatting. */
    public const READ_DATA_ONLY = 1;

    /** Do not read charts. */
    public const SKIP_EMPTY_CELLS = 2;

    /** Ignore the row/column dimension records. */
    public const IGNORE_ROWS_WITH_NO_CELLS = 4;

    /** True when this reader can open the given file. */
    public function canRead(string $filename): bool;

    /** Load a workbook from a path or stream-wrapper URL. */
    public function load(string $filename, int $flags = 0): Spreadsheet;
}
