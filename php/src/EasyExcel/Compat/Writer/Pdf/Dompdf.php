<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer\Pdf;

use EasyExcel\Compat\Writer\Exception;
use EasyExcel\Compat\Writer\Pdf;

/**
 * PDF driver backed by dompdf/dompdf, which the consumer requires.
 */
class Dompdf extends Pdf
{
    /** The external writer. Overridable so a test can inject a double. */
    protected function createExternalWriterInstance(): object
    {
        if (!\class_exists(\Dompdf\Dompdf::class)) {
            throw new Exception('dompdf/dompdf is required for the Dompdf writer; run composer require dompdf/dompdf');
        }

        return new \Dompdf\Dompdf();
    }

    /**
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        $pdf = $this->createExternalWriterInstance();
        $pdf->setPaper($this->paperSizeForDompdf(), $this->resolveOrientation());
        $pdf->loadHtml($this->generateHtmlAll());
        $pdf->render();

        $fileHandle = $this->prepareForSave($filename);
        \fwrite($fileHandle, (string) $pdf->output());
        $this->restoreStateAfterSave($fileHandle);
    }

    /**
     * Dompdf takes either a size name or a full [x0, y0, x1, y1] bounding box,
     * so a size that resolves to a bare [width, height] pair has to be
     * expanded rather than passed through.
     *
     * @return float[]|string
     */
    private function paperSizeForDompdf(): array|string
    {
        $size = $this->resolvePaperSize();

        if (\is_array($size) && \count($size) === 2) {
            return [0.0, 0.0, $size[0], $size[1]];
        }

        return $size;
    }
}
