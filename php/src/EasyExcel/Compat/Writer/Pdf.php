<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;

use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\PageSetup;
use EasyExcel\Compat\Worksheet\Worksheet;

/**
 * Base for the PDF writers: the HTML renderer plus the page setup, with the
 * HTML->PDF step left to a driver subclass.
 *
 * Mirrors PhpSpreadsheet's own shape — `Writer\Pdf` is abstract there too, and
 * the consumer picks a driver by class:
 *
 *     $writer = new Writer\Pdf\Mpdf($spreadsheet);       // or Tcpdf, Dompdf
 *     $writer->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
 *     $writer->save('report.pdf');
 *
 * Each driver depends on its own library (mpdf/mpdf, tecnickcom/tcpdf,
 * dompdf/dompdf), which the consumer requires; instantiating a driver whose
 * library is absent throws. `Pdf\Snappy` is the exception — it shells out to
 * wkhtmltopdf through knplabs/knp-snappy rather than embedding an engine, and
 * has no counterpart upstream.
 */
abstract class Pdf extends Html
{
    /**
     * Paper size names, keyed by the PageSetup::PAPERSIZE_* constant.
     *
     * Mirrors PhpSpreadsheet's own xRef list so that callers reading it back
     * (or subclasses overriding it) see the values they expect. Sizes with no
     * standard name carry their [width, height] in points.
     *
     * @var array<int, float[]|string>
     */
    protected static array $paperSizes = [
        PageSetup::PAPERSIZE_LETTER => 'LETTER',
        PageSetup::PAPERSIZE_LETTER_SMALL => 'LETTER',
        PageSetup::PAPERSIZE_TABLOID => [792.00, 1224.00],
        PageSetup::PAPERSIZE_LEDGER => [1224.00, 792.00],
        PageSetup::PAPERSIZE_LEGAL => 'LEGAL',
        PageSetup::PAPERSIZE_STATEMENT => [396.00, 612.00],
        PageSetup::PAPERSIZE_EXECUTIVE => 'EXECUTIVE',
        PageSetup::PAPERSIZE_A3 => 'A3',
        PageSetup::PAPERSIZE_A4 => 'A4',
        PageSetup::PAPERSIZE_A4_SMALL => 'A4',
        PageSetup::PAPERSIZE_A5 => 'A5',
        PageSetup::PAPERSIZE_B4 => 'B4',
        PageSetup::PAPERSIZE_B5 => 'B5',
        PageSetup::PAPERSIZE_FOLIO => 'FOLIO',
        PageSetup::PAPERSIZE_QUARTO => [609.45, 779.53],
        PageSetup::PAPERSIZE_STANDARD_1 => [720.00, 1008.00],
        PageSetup::PAPERSIZE_STANDARD_2 => [792.00, 1224.00],
        PageSetup::PAPERSIZE_NOTE => 'LETTER',
    ];

    /** Temporary storage directory. */
    protected string $tempDir = '';

    /** Font. */
    protected string $font = 'freesans';

    /** Orientation override; null means "use the sheet's own page setup". */
    protected ?string $orientation = null;

    /** Paper size override; null means "use the sheet's own page setup". */
    protected ?int $paperSize = null;

    /**
     * The workbook being written.
     *
     * Html keeps its own copy private, so hold a second reference rather than
     * widening the parent's property.
     */
    private Spreadsheet $spreadsheetForPdf;

    public function __construct(Spreadsheet $spreadsheet)
    {
        parent::__construct($spreadsheet);
        $this->spreadsheetForPdf = $spreadsheet;
        $this->tempDir = \sys_get_temp_dir() . '/phpsppdf';
    }

    /**
     * The workbook being written.
     *
     * Html keeps its own reference private, so subclasses that need the
     * workbook — an app writer wrapping this one to drive an external
     * converter, say — would otherwise have to hold a third copy.
     */
    /**
     * Write the PDF. Drivers override this.
     *
     * PHP will not let an inherited concrete method be re-declared abstract,
     * so the base cannot force the override at compile time; refuse at run
     * time instead, rather than letting Html::save() quietly write an HTML
     * document into a .pdf file.
     *
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        throw new Exception(static::class . ' must implement save(); the base writes HTML, not PDF.');
    }

    public function getSpreadsheet(): Spreadsheet
    {
        return $this->spreadsheetForPdf;
    }

    public function getFont(): string
    {
        return $this->font;
    }

    /**
     * Set the font. Valid values are the ones the downstream renderer knows;
     * they are stored verbatim and never interpreted here.
     */
    public function setFont(string $fontName): static
    {
        $this->font = $fontName;

        return $this;
    }

    /** The paper-size override, or null when the sheet's page setup wins. */
    public function getPaperSize(): ?int
    {
        return $this->paperSize;
    }

    public function setPaperSize(int $paperSize): static
    {
        $this->paperSize = $paperSize;

        return $this;
    }

    /** The orientation override, or null when the sheet's page setup wins. */
    public function getOrientation(): ?string
    {
        return $this->orientation;
    }

    public function setOrientation(string $orientation): static
    {
        $this->orientation = $orientation;

        return $this;
    }

    public function getTempDir(): string
    {
        return $this->tempDir;
    }

    public function setTempDir(string $temporaryDirectory): static
    {
        $this->tempDir = $temporaryDirectory;

        return $this;
    }

    /**
     * The paper size that should be used, as a name ('A4', 'LETTER', ...) or a
     * [width, height] pair in points.
     *
     * Resolves the override first, then the active sheet's page setup, then
     * PhpSpreadsheet's own default of Letter. Handy for building the argument
     * list for an external converter.
     *
     * @return float[]|string
     */
    public function resolvePaperSize(): array|string
    {
        $size = $this->paperSize ?? $this->activeSheetPageSetup()->getPaperSize();

        return static::$paperSizes[$size] ?? 'LETTER';
    }

    /**
     * The orientation that should be used: 'portrait' or 'landscape'.
     *
     * Resolves the override first, then the active sheet's page setup. A sheet
     * left on ORIENTATION_DEFAULT reports portrait, matching what every PDF
     * engine assumes when no orientation is given.
     */
    public function resolveOrientation(): string
    {
        $orientation = $this->orientation ?? $this->activeSheetPageSetup()->getOrientation();

        return $orientation === PageSetup::ORIENTATION_LANDSCAPE
            ? PageSetup::ORIENTATION_LANDSCAPE
            : PageSetup::ORIENTATION_PORTRAIT;
    }

    /**
     * Open the target for writing and hand back the handle.
     *
     * Drivers that build the PDF in memory can ignore this and write the
     * bytes themselves; it exists for the ones that stream.
     *
     * @param resource|string $filename
     *
     * @return resource
     */
    protected function prepareForSave($filename)
    {
        $fileHandle = \fopen($filename, 'w');
        if ($fileHandle === false) {
            throw new Exception("Could not open file $filename for writing.");
        }

        return $fileHandle;
    }

    /** @param resource $fileHandle */
    protected function restoreStateAfterSave($fileHandle): void
    {
        if (\is_resource($fileHandle)) {
            \fclose($fileHandle);
        }
    }

    /** Millimetres per inch. */
    private const MM_PER_INCH = 25.4;

    /** PostScript points per inch. */
    private const POINTS_PER_INCH = 72.0;

    /**
     * The page margins in inches, as the worksheet holds them.
     *
     * A worksheet that never had a margin set reports -1 (the "unset"
     * sentinel), which every engine rejects; those fall back to
     * PhpSpreadsheet's own defaults.
     *
     * @return array{left: float, right: float, top: float, bottom: float}
     */
    protected function marginsInInches(): array
    {
        $margins = $this->activeSheet()->getPageMargins();

        return [
            'left' => self::marginOrDefault($margins->getLeft(), 0.7),
            'right' => self::marginOrDefault($margins->getRight(), 0.7),
            'top' => self::marginOrDefault($margins->getTop(), 0.75),
            'bottom' => self::marginOrDefault($margins->getBottom(), 0.75),
        ];
    }

    /**
     * The page margins in millimetres, for engines measuring in mm.
     *
     * @return array{left: float, right: float, top: float, bottom: float}
     */
    protected function marginsInMm(): array
    {
        return \array_map(
            static fn (float $inches): float => $inches * self::MM_PER_INCH,
            $this->marginsInInches()
        );
    }

    /**
     * The page margins in PostScript points, for engines measuring in pt.
     *
     * @return array{left: float, right: float, top: float, bottom: float}
     */
    protected function marginsInPoints(): array
    {
        return \array_map(
            static fn (float $inches): float => $inches * self::POINTS_PER_INCH,
            $this->marginsInInches()
        );
    }

    private static function marginOrDefault(float $inches, float $default): float
    {
        return $inches < 0 ? $default : $inches;
    }

    /** Convert inches to millimetres. */
    protected function inchesToMm(float $inches): float
    {
        return $inches * self::MM_PER_INCH;
    }

    protected function activeSheet(): Worksheet
    {
        $index = $this->getSheetIndex();

        return $index === null
            ? $this->spreadsheetForPdf->getActiveSheet()
            : $this->spreadsheetForPdf->getSheet($index);
    }

    private function activeSheetPageSetup(): PageSetup
    {
        return $this->activeSheet()->getPageSetup();
    }
}
