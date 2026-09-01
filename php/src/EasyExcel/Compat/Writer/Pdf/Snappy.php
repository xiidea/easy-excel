<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer\Pdf;

use EasyExcel\Compat\Writer\Exception;
use EasyExcel\Compat\Writer\Pdf;

/**
 * PDF driver backed by knplabs/knp-snappy, which shells out to wkhtmltopdf.
 *
 * No counterpart upstream: the others embed a PHP rendering engine, this one
 * drives an external binary. Because the binary renders with a real browser
 * engine it handles CSS that the PHP engines do not, which is why an app with
 * heavily styled reports tends to reach for it.
 *
 * The Snappy instance is injected rather than constructed, since it carries
 * the binary path and per-app defaults:
 *
 *     $writer = new Writer\Pdf\Snappy($spreadsheet);
 *     $writer->setSnappy($snappy);
 *     $writer->save('report.pdf');
 */
class Snappy extends Pdf
{
    private ?object $snappy = null;

    /**
     * @param object $snappy a \Knp\Snappy\Pdf (typed loosely so the polyfill
     *                       does not hard-depend on knplabs/knp-snappy)
     */
    public function setSnappy(object $snappy): static
    {
        $this->snappy = $snappy;

        return $this;
    }

    public function getSnappy(): ?object
    {
        return $this->snappy;
    }

    /**
     * The wkhtmltopdf options derived from the workbook's page setup.
     *
     * Exposed so a caller can inspect or extend them before saving.
     *
     * @return array<string, float|string>
     */
    public function getPageOptions(): array
    {
        $paperSize = $this->resolvePaperSize();
        $margins = $this->marginsInMm();

        return [
            'page-size' => \is_string($paperSize) ? $paperSize : 'LETTER',
            'orientation' => \strtoupper($this->resolveOrientation()),
            'margin-left' => $margins['left'],
            'margin-right' => $margins['right'],
            'margin-top' => $margins['top'],
            'margin-bottom' => $margins['bottom'],
            'title' => (string) $this->getSpreadsheet()->getProperties()->getTitle(),
        ];
    }

    /**
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        if ($this->snappy === null) {
            throw new Exception('No Snappy instance set; call setSnappy() with a \Knp\Snappy\Pdf first.');
        }

        foreach ($this->getPageOptions() as $option => $value) {
            $this->snappy->setOption($option, $value);
        }

        $fileHandle = $this->prepareForSave($filename);
        \fwrite($fileHandle, (string) $this->snappy->getOutputFromHtml($this->generateHtmlAll()));
        $this->restoreStateAfterSave($fileHandle);
    }
}
