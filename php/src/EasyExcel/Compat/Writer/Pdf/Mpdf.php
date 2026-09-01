<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer\Pdf;

use EasyExcel\Compat\Worksheet\PageSetup;
use EasyExcel\Compat\Writer\Exception;
use EasyExcel\Compat\Writer\Pdf;

/**
 * PDF driver backed by mpdf/mpdf, which the consumer requires.
 */
class Mpdf extends Pdf
{
    /**
     * The external writer.
     *
     * Returns `object` rather than the concrete class so the hook stays
     * usable when the library is absent — a subclass may inject a double.
     *
     * @param mixed[] $config
     */
    protected function createExternalWriterInstance(array $config): object
    {
        if (!\class_exists(\Mpdf\Mpdf::class)) {
            throw new Exception('mpdf/mpdf is required for the Mpdf writer; run composer require mpdf/mpdf');
        }

        return new \Mpdf\Mpdf($config);
    }

    /**
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        $margins = $this->marginsInMm();
        $pdf = $this->createExternalWriterInstance(['tempDir' => $this->getTempDir() . '/mpdf']);

        // _setPageSize() takes its orientation by reference and mutates it,
        // so it needs a variable rather than an expression.
        $orientation = $this->orientationLetter();
        $orientationForPageSize = $orientation;
        $pdf->_setPageSize($this->paperSizeName(), $orientationForPageSize);
        $pdf->DefOrientation = $orientation;
        $pdf->AddPageByArray([
            'orientation' => $orientation,
            'margin-left' => $margins['left'],
            'margin-right' => $margins['right'],
            'margin-top' => $margins['top'],
            'margin-bottom' => $margins['bottom'],
        ]);

        $properties = $this->getSpreadsheet()->getProperties();
        $pdf->SetTitle($properties->getTitle());
        $pdf->SetAuthor($properties->getCreator());
        $pdf->SetSubject($properties->getSubject());
        $pdf->SetCreator($properties->getCreator());
        // Keywords are not part of the Compat Properties surface, so there is
        // nothing to forward; mPDF simply leaves the field empty.

        $pdf->WriteHTML($this->generateHtmlAll());

        $fileHandle = $this->prepareForSave($filename);
        \fwrite($fileHandle, $pdf->Output('', 'S'));
        $this->restoreStateAfterSave($fileHandle);
    }

    /** mPDF wants 'L' or 'P'. */
    private function orientationLetter(): string
    {
        return $this->resolveOrientation() === PageSetup::ORIENTATION_LANDSCAPE ? 'L' : 'P';
    }

    /** mPDF wants a named size; anonymous point sizes fall back to A4. */
    private function paperSizeName(): string
    {
        $size = $this->resolvePaperSize();

        return \is_string($size) ? $size : 'A4';
    }
}
