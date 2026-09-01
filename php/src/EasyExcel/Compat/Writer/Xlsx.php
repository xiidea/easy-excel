<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;

use EasyExcel\Compat\Shared\StreamPath;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Native;

class Xlsx extends BaseWriter
{
    private string $password = '';

    public function __construct(private Spreadsheet $spreadsheet)
    {
    }

    /**
     * easy-excel extra (PhpSpreadsheet cannot write encrypted xlsx): a
     * non-empty password produces an agile-encrypted container. Encryption
     * routes streamed auto-filters through the save-time degrade.
     */
    public function setPassword(string $password): static
    {
        $this->password = $password;

        return $this;
    }

    /**
     * Saves to a filesystem path, a stream-wrapper URL (php://, gaufrette://,
     * ...), or an open resource. Streams go through a temp file because the
     * extension writes files directly (the xlsx container is already deflated
     * — never double-compress it, PLAN.md B10).
     *
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        $this->processFlags($flags);
        $this->spreadsheet->flushAll();
        $handle = $this->spreadsheet->getHandle();

        // Opt-in only. PhpSpreadsheet pre-calculates by default, but doing so
        // here reads every formula back and forces a streamed workbook into
        // the full in-memory model — an OOM risk on million-row exports. The
        // getter still reports upstream's default; only an explicit
        // setPreCalculateFormulas() call enables the pass (COMPAT.md §24).
        Native::setPrecalculateFormulas($handle, $this->shouldPreCalculateFormulas());

        if (\is_string($filename) && !StreamPath::isWrapped($filename)) {
            Native::saveXlsx($handle, $filename, $this->password);

            return;
        }

        $tmp = \tempnam(\sys_get_temp_dir(), 'eexcel');
        if ($tmp === false) {
            throw new Exception('Could not create temporary file');
        }
        // excelize's SaveAs validates the file extension and tempnam()
        // produces none — stage under a .xlsx name or the native save fails
        // with "unsupported workbook file format".
        $staged = $tmp . '.xlsx';
        try {
            Native::saveXlsx($handle, $staged, $this->password);
            $this->openFileHandle($filename);
            $in = \fopen($staged, 'rb');
            \stream_copy_to_stream($in, $this->fileHandle);
            \fclose($in);
            $this->maybeCloseFileHandle();
        } finally {
            @\unlink($staged);
            @\unlink($tmp);
        }
    }
}
