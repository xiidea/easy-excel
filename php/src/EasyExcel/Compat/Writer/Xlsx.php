<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;

use EasyExcel\Compat\Exception;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Native;

class Xlsx
{
    public function __construct(private Spreadsheet $spreadsheet)
    {
    }

    /**
     * Saves to a filesystem path or a php:// stream. Streams go through a
     * temp file because the extension writes files directly (the xlsx
     * container is already deflated — never double-compress it, PLAN.md B10).
     */
    public function save(string $filename, int $flags = 0): void
    {
        $this->spreadsheet->flushAll();
        $handle = $this->spreadsheet->getHandle();

        if (!\str_starts_with($filename, 'php://')) {
            Native::saveXlsx($handle, $filename);

            return;
        }

        $tmp = \tempnam(\sys_get_temp_dir(), 'eexcel');
        if ($tmp === false) {
            throw new Exception('Could not create temporary file');
        }
        try {
            Native::saveXlsx($handle, $tmp);
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
