<?php

declare(strict_types=1);

/*
 * Writer\Pdf is an abstract base carrying the HTML renderer and the page
 * setup; each driver subclass supplies the HTML->PDF step. Mirrors
 * PhpSpreadsheet, where the consumer picks a driver by class rather than
 * through IOFactory.
 *
 * The bundled drivers depend on libraries the polyfill does not require
 * (mpdf, tcpdf, dompdf, knp-snappy), so the tests cover the base contract and
 * the driver plumbing through doubles rather than rendering real PDFs.
 */

/** Snappy stand-in: records the options and returns identifiable bytes. */
final class FakeSnappy
{
    /** @var array<string, mixed> */
    public array $options = [];

    public string $html = '';

    public function setOption(string $name, mixed $value): void
    {
        $this->options[$name] = $value;
    }

    public function getOutputFromHtml(string $html): string
    {
        $this->html = $html;

        return "%PDF-1.4 fake(" . \strlen($html) . ")";
    }
}

return [
    'pdf: the base is abstract' => function (): void {
        $reflection = new ReflectionClass(PhpOffice\PhpSpreadsheet\Writer\Pdf::class);

        T::ok($reflection->isAbstract(), 'Writer\Pdf must not be instantiable on its own');
    },

    'pdf: IOFactory points at the drivers instead of building one' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();

        // IOFactory must never hand back the abstract base: either it refuses
        // and names a driver, or another case registered a concrete one for
        // 'Pdf' (registerWriter state is static).
        try {
            $writer = PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Pdf');
            T::ok(
                $writer instanceof PhpOffice\PhpSpreadsheet\Writer\Pdf,
                'a registered driver came back'
            );
            T::ok(!(new ReflectionClass($writer))->isAbstract(), 'and it is concrete');
        } catch (\Throwable $e) {
            T::ok(\str_contains($e->getMessage(), 'abstract'), 'says why it refused');
            T::ok(\str_contains($e->getMessage(), 'Pdf\Mpdf'), 'names a driver');
        }
    },

    'pdf: every bundled driver extends the base' => function (): void {
        foreach (['Mpdf', 'Tcpdf', 'Dompdf', 'Snappy'] as $driver) {
            $class = 'PhpOffice\PhpSpreadsheet\Writer\Pdf\\' . $driver;
            T::ok(\class_exists($class), "$driver resolves through the alias");
            T::ok(
                \is_subclass_of($class, PhpOffice\PhpSpreadsheet\Writer\Pdf::class),
                "$driver extends Writer\Pdf"
            );
        }
    },

    'pdf: the page-setup accessors round-trip on a driver' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);

        T::same(null, $writer->getPaperSize(), 'no override until one is set');
        T::same(null, $writer->getOrientation(), 'no override until one is set');
        T::same('freesans', $writer->getFont());

        $writer->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
        $writer->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
        $writer->setFont('dejavusans');
        $writer->setTempDir('/tmp/pdf-writer-test');

        T::same(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4, $writer->getPaperSize());
        T::same('landscape', $writer->getOrientation());
        T::same('dejavusans', $writer->getFont());
        T::same('/tmp/pdf-writer-test', $writer->getTempDir());
        T::ok($writer->getSpreadsheet() === $spreadsheet, 'the workbook is reachable');
    },

    'pdf: an override wins over the sheet page setup' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $spreadsheet->getActiveSheet()->getPageSetup()
            ->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_LEGAL)
            ->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_PORTRAIT);
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);

        T::same('LEGAL', $writer->resolvePaperSize(), 'the sheet wins while no override is set');
        T::same('portrait', $writer->resolveOrientation());

        $writer->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A3);
        $writer->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);

        T::same('A3', $writer->resolvePaperSize(), 'the override wins once set');
        T::same('landscape', $writer->resolveOrientation());
    },

    'pdf: a sheet left on the default orientation resolves to portrait' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);

        T::same('portrait', $writer->resolveOrientation(), 'every engine assumes portrait when unset');
    },

    'pdf: a paper size with no standard name resolves to its point size' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);
        $writer->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_TABLOID);

        T::same([792.00, 1224.00], $writer->resolvePaperSize());
    },

    'pdf: unset margins fall back rather than going negative' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'x');
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);

        $options = $writer->getPageOptions();

        foreach (['margin-left', 'margin-right', 'margin-top', 'margin-bottom'] as $margin) {
            T::ok($options[$margin] > 0, "$margin must be positive, not the -1 sentinel");
        }
        T::same(17.78, \round((float) $options['margin-left'], 2), '0.7in in mm');
        T::same(19.05, \round((float) $options['margin-top'], 2), '0.75in in mm');
    },

    'pdf: the Snappy driver drives the injected instance' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $spreadsheet->getActiveSheet()->setCellValue('A1', 'Trial Balance');
        $spreadsheet->getActiveSheet()->getPageSetup()
            ->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);

        $snappy = new FakeSnappy();
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);
        $writer->setSnappy($snappy);

        $target = \sys_get_temp_dir() . '/easy-excel-snappy-' . \bin2hex(\random_bytes(4)) . '.pdf';
        $writer->save($target);

        T::same('LANDSCAPE', $snappy->options['orientation'], 'orientation reaches the driver');
        T::same('LETTER', $snappy->options['page-size']);
        T::ok(\str_contains($snappy->html, 'Trial Balance'), 'the rendered HTML is handed over');
        T::ok(\str_starts_with((string) \file_get_contents($target), '%PDF-'), 'the bytes are written');
        @\unlink($target);
    },

    'pdf: the Snappy driver refuses to save without an instance' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Pdf\Snappy($spreadsheet);

        T::throws(
            \EasyExcel\Compat\Writer\Exception::class,
            static fn () => $writer->save(\sys_get_temp_dir() . '/never-written.pdf'),
            'a missing Snappy instance must fail loudly'
        );
    },

    'pdf: a driver whose library is absent says which package to install' => function (): void {
        $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();

        foreach ([['Mpdf', 'mpdf/mpdf'], ['Tcpdf', 'tecnickcom/tcpdf'], ['Dompdf', 'dompdf/dompdf']] as [$driver, $package]) {
            $class = 'PhpOffice\PhpSpreadsheet\Writer\Pdf\\' . $driver;
            $writer = new $class($spreadsheet);
            try {
                $writer->save(\sys_get_temp_dir() . '/never-written.pdf');
                T::ok(false, "$driver saved without its library present");
            } catch (\Throwable $e) {
                T::ok(\str_contains($e->getMessage(), $package), "$driver names $package");
            }
        }
    },
];
