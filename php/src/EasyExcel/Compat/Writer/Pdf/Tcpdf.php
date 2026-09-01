<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer\Pdf;

use EasyExcel\Compat\Worksheet\PageSetup;
use EasyExcel\Compat\Writer\Exception;
use EasyExcel\Compat\Writer\Pdf;

/**
 * PDF driver backed by tecnickcom/tcpdf, which the consumer requires.
 */
class Tcpdf extends Pdf
{
    /**
     * The external writer.
     *
     * Returns `object` rather than the concrete class so the hook stays
     * usable when the library is absent — a subclass may inject a double.
     *
     * @param float[]|string $paperSize
     */
    protected function createExternalWriterInstance(string $orientation, string $unit, $paperSize): object
    {
        if (!\class_exists(\TCPDF::class)) {
            throw new Exception('tecnickcom/tcpdf is required for the Tcpdf writer; run composer require tecnickcom/tcpdf');
        }

        return new \TCPDF($orientation, $unit, $paperSize);
    }

    /**
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        $orientation = $this->resolveOrientation() === PageSetup::ORIENTATION_LANDSCAPE ? 'L' : 'P';
        // TCPDF is constructed with unit 'pt', so the margins must be points.
        $margins = $this->marginsInPoints();

        $pdf = $this->createExternalWriterInstance($orientation, 'pt', $this->resolvePaperSize());
        $pdf->setFontSubsetting(false);
        $pdf->SetTitle($this->getSpreadsheet()->getProperties()->getTitle());
        $pdf->setPrintHeader(false);
        $pdf->setPrintFooter(false);
        $pdf->AddPage();
        $pdf->SetFont($this->getFont());
        $pdf->SetMargins($margins['left'], $margins['top'], $margins['right']);
        $pdf->SetAutoPageBreak(true, $margins['bottom']);

        $pdf->writeHTML($this->generateHtmlAll());

        $fileHandle = $this->prepareForSave($filename);
        \fwrite($fileHandle, $pdf->Output('', 'S'));
        $this->restoreStateAfterSave($fileHandle);
    }
}
