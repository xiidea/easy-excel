<?php

declare(strict_types=1);

/*
 * Driver save paths.
 *
 * The bundled drivers each depend on a library the polyfill does not require,
 * so these tests inject a double through createExternalWriterInstance() and
 * assert what the driver hands the engine: units, orientation, paper size and
 * the rendered HTML. Without this the drivers were only ever exercised up to
 * their "library missing" guard, which sits before every line that matters.
 */

/** Records what a driver calls, standing in for mPDF/TCPDF/Dompdf. */
final class FakePdfEngine
{
    /** @var array<int, array{0: string, 1: array<int, mixed>}> */
    public array $calls = [];

    public string $html = '';

    /** mPDF writes page setup onto a public property. */
    public string $DefOrientation = '';

    public function __call(string $method, array $arguments): mixed
    {
        $this->calls[] = [$method, $arguments];

        return null;
    }

    /** mPDF mutates the orientation it is handed, hence the reference. */
    public function _setPageSize(mixed $size, mixed &$orientation): void
    {
        $this->calls[] = ['_setPageSize', [$size, $orientation]];
    }

    /** mPDF and TCPDF differ only in case, which PHP does not distinguish. */
    public function WriteHTML(string $html): void
    {
        $this->html = $html;
    }

    public function loadHtml(string $html): void
    {
        $this->html = $html;
    }

    /**
     * mPDF/TCPDF call Output($name, $dest); Dompdf calls output(). PHP does
     * not distinguish the case, so one method serves both.
     */
    public function Output(string $name = '', string $dest = ''): string
    {
        return '%PDF-1.4 fake-engine';
    }

    /** @return array<int, mixed>|null arguments of the first $method call */
    public function firstCall(string $method): ?array
    {
        foreach ($this->calls as [$name, $arguments]) {
            if ($name === $method) {
                return $arguments;
            }
        }

        return null;
    }
}

final class TestableMpdf extends EasyExcel\Compat\Writer\Pdf\Mpdf
{
    public FakePdfEngine $engine;

    protected function createExternalWriterInstance(array $config): object
    {
        return $this->engine ??= new FakePdfEngine();
    }
}

final class TestableTcpdf extends EasyExcel\Compat\Writer\Pdf\Tcpdf
{
    public FakePdfEngine $engine;

    /** @var array<int, mixed> */
    public array $constructorArgs = [];

    protected function createExternalWriterInstance(string $orientation, string $unit, $paperSize): object
    {
        $this->constructorArgs = [$orientation, $unit, $paperSize];

        return $this->engine ??= new FakePdfEngine();
    }
}

final class TestableDompdf extends EasyExcel\Compat\Writer\Pdf\Dompdf
{
    public FakePdfEngine $engine;

    protected function createExternalWriterInstance(): object
    {
        return $this->engine ??= new FakePdfEngine();
    }
}

/** @return array{0: PhpOffice\PhpSpreadsheet\Spreadsheet, 1: string} */
function pdfDriverFixture(): array
{
    $spreadsheet = new PhpOffice\PhpSpreadsheet\Spreadsheet();
    $spreadsheet->getActiveSheet()->setCellValue('A1', 'Trial Balance');

    return [$spreadsheet, \sys_get_temp_dir() . '/easy-excel-driver-' . \bin2hex(\random_bytes(4)) . '.pdf'];
}

return [
    'driver: Tcpdf measures margins in the unit it constructs with' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $writer = new TestableTcpdf($spreadsheet);

        $writer->save($target);

        [, $unit] = $writer->constructorArgs;
        T::same('pt', $unit, 'the document is built in points');

        // 0.7in left / 0.75in top, in points, not the 17.78/19.05 of millimetres.
        $margins = $writer->engine->firstCall('SetMargins');
        T::same(50.4, \round((float) $margins[0], 1), 'left margin in points');
        T::same(54.0, \round((float) $margins[1], 1), 'top margin in points');
        T::same(50.4, \round((float) $margins[2], 1), 'right margin in points');

        $pageBreak = $writer->engine->firstCall('SetAutoPageBreak');
        T::same(54.0, \round((float) $pageBreak[1], 1), 'bottom margin in points');
        @\unlink($target);
    },

    'driver: Tcpdf forwards orientation and the rendered HTML' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $spreadsheet->getActiveSheet()->getPageSetup()
            ->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
        $writer = new TestableTcpdf($spreadsheet);

        $writer->save($target);

        T::same('L', $writer->constructorArgs[0], 'landscape reaches the engine');
        T::ok(\str_contains($writer->engine->html, 'Trial Balance'), 'the sheet is rendered');
        T::ok(\str_starts_with((string) \file_get_contents($target), '%PDF-'), 'bytes are written');
        @\unlink($target);
    },

    'driver: Mpdf survives the by-reference page-size call' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $spreadsheet->getActiveSheet()->getPageSetup()
            ->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
        $writer = new TestableMpdf($spreadsheet);

        // _setPageSize() takes its second argument by reference; passing an
        // expression there is a fatal, so reaching this line at all is the
        // assertion.
        $writer->save($target);

        $pageSize = $writer->engine->firstCall('_setPageSize');
        T::same('LETTER', $pageSize[0], 'the paper name is passed');
        T::same('L', $pageSize[1], 'the orientation is passed');
        T::same('L', $writer->engine->DefOrientation, 'and recorded on the engine');
        @\unlink($target);
    },

    'driver: Mpdf forwards only the properties Compat exposes' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $writer = new TestableMpdf($spreadsheet);

        $writer->save($target);

        $called = \array_column($writer->engine->calls, 0);
        foreach (['SetTitle', 'SetAuthor', 'SetSubject', 'SetCreator'] as $setter) {
            T::ok(\in_array($setter, $called, true), "$setter is forwarded");
        }
        T::ok(!\in_array('SetKeywords', $called, true), 'no keywords: Compat has no getKeywords()');
        T::ok(\str_contains($writer->engine->html, 'Trial Balance'), 'the sheet is rendered');
        @\unlink($target);
    },

    'driver: Mpdf falls back to A4 for a size with no name' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $writer = new TestableMpdf($spreadsheet);
        $writer->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_TABLOID);

        $writer->save($target);

        T::same('A4', $writer->engine->firstCall('_setPageSize')[0], 'mPDF needs a named size');
        @\unlink($target);
    },

    'driver: Dompdf expands a bare size pair into a bounding box' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $writer = new TestableDompdf($spreadsheet);
        $writer->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_TABLOID);

        $writer->save($target);

        $paper = $writer->engine->firstCall('setPaper');
        T::same([0.0, 0.0, 792.00, 1224.00], $paper[0], 'Dompdf needs four elements, not two');
        T::same('portrait', $paper[1]);
        @\unlink($target);
    },

    'driver: Dompdf passes a named size straight through' => function (): void {
        [$spreadsheet, $target] = pdfDriverFixture();
        $writer = new TestableDompdf($spreadsheet);
        $writer->setPaperSize(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4);
        $writer->setOrientation(PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);

        $writer->save($target);

        $paper = $writer->engine->firstCall('setPaper');
        T::same('A4', $paper[0], 'a named size is not expanded');
        T::same('landscape', $paper[1]);
        T::ok(\str_starts_with((string) \file_get_contents($target), '%PDF-'), 'bytes are written');
        @\unlink($target);
    },

    'driver: every driver declares its own save()' => function (): void {
        foreach (['Mpdf', 'Tcpdf', 'Dompdf', 'Snappy'] as $driver) {
            $class = 'EasyExcel\Compat\Writer\Pdf\\' . $driver;
            $method = new ReflectionMethod($class, 'save');
            T::same(
                $class,
                $method->getDeclaringClass()->getName(),
                "$driver must not inherit a save() that writes HTML"
            );
        }
    },
];
