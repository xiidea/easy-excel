<?php

declare(strict_types=1);

use EasyExcel\Compat\Cell\AddressRange;
use EasyExcel\Compat\Cell\CellAddress;
use EasyExcel\Compat\Exception as CompatException;
use EasyExcel\Compat\Settings;
use EasyExcel\Compat\Shared\Drawing as SharedDrawing;
use EasyExcel\Compat\Shared\File as SharedFile;
use EasyExcel\Compat\Shared\Font as SharedFont;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\BaseDrawing;
use EasyExcel\Compat\Worksheet\Drawing;
use EasyExcel\Compat\Worksheet\MemoryDrawing;

/**
 * Compat's Style\Font is bound to its owning Style and cannot be constructed
 * standalone, so the Shared helpers accept any object exposing getName()/
 * getSize(). This is the minimal such object — it also proves the duck-typed
 * contract holds for callers that are not Compat's own Font.
 */
final class FontStub
{
    public function __construct(private string $name = 'Calibri', private float $size = 11.0)
    {
    }

    public function getName(): string
    {
        return $this->name;
    }

    public function getSize(): float
    {
        return $this->size;
    }
}

return [
    // ---------------------------------------------------------------- exceptions

    'wave52: Writer\Exception is catchable both narrowly and broadly' => function (): void {
        $e = new \EasyExcel\Compat\Writer\Exception('boom');
        T::ok($e instanceof \EasyExcel\Compat\Writer\Exception, 'is the narrow type');
        T::ok($e instanceof CompatException, 'still the flat Compat exception');
        T::ok($e instanceof \RuntimeException, 'still a RuntimeException');
        T::same('boom', $e->getMessage());
    },

    'wave52: Reader\Exception is catchable both narrowly and broadly' => function (): void {
        $e = new \EasyExcel\Compat\Reader\Exception('nope');
        T::ok($e instanceof \EasyExcel\Compat\Reader\Exception, 'is the narrow type');
        T::ok($e instanceof CompatException, 'still the flat Compat exception');
        T::same('nope', $e->getMessage());
    },

    'wave52: Calculation\Exception is catchable both narrowly and broadly' => function (): void {
        $e = new \EasyExcel\Compat\Calculation\Exception('bad formula');
        T::ok($e instanceof \EasyExcel\Compat\Calculation\Exception, 'is the narrow type');
        T::ok($e instanceof CompatException, 'still the flat Compat exception');
    },

    'wave52: the three exception types are siblings, not each other' => function (): void {
        $w = new \EasyExcel\Compat\Writer\Exception('w');
        T::ok(!($w instanceof \EasyExcel\Compat\Reader\Exception), 'writer is not a reader exception');
        T::ok(!($w instanceof \EasyExcel\Compat\Calculation\Exception), 'writer is not a calc exception');
    },

    'wave52: writers actually throw the narrow Writer\Exception' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $w = new \EasyExcel\Compat\Writer\Csv($s);
        // Only '"' is a supported enclosure (COMPAT.md divergence 5).
        T::throws(
            \EasyExcel\Compat\Writer\Exception::class,
            static fn () => $w->setEnclosure("'"),
            'bad enclosure throws the narrow writer type',
        );
    },

    'wave52: readers actually throw the narrow Reader\Exception' => function (): void {
        EasyExcelFake::reset();
        $r = new \EasyExcel\Compat\Reader\Xlsx();
        T::throws(
            \EasyExcel\Compat\Reader\Exception::class,
            static fn () => $r->load('/nonexistent/easy-excel/nope.xlsx'),
            'missing file throws the narrow reader type',
        );
    },

    'wave52: a broad catch still catches the narrowed writer throw' => function (): void {
        EasyExcelFake::reset();
        $w = new \EasyExcel\Compat\Writer\Csv(new Spreadsheet());
        $caught = null;
        try {
            $w->setEnclosure('|');
        } catch (CompatException $e) {
            $caught = $e;
        }
        T::ok($caught !== null, 'existing broad catch blocks keep working');
    },

    // ------------------------------------------------------------------- IReader

    'wave52: both readers implement IReader' => function (): void {
        T::ok(new \EasyExcel\Compat\Reader\Xlsx() instanceof \EasyExcel\Compat\Reader\IReader, 'Xlsx');
        T::ok(new \EasyExcel\Compat\Reader\Csv() instanceof \EasyExcel\Compat\Reader\IReader, 'Csv');
    },

    'wave52: IReader satisfies a typed parameter' => function (): void {
        $take = static fn (\EasyExcel\Compat\Reader\IReader $r): string => $r::class;
        T::same(\EasyExcel\Compat\Reader\Csv::class, $take(new \EasyExcel\Compat\Reader\Csv()));
    },

    'wave52: IReader exposes the documented flag constants' => function (): void {
        T::same(1, \EasyExcel\Compat\Reader\IReader::READ_DATA_ONLY);
        T::same(2, \EasyExcel\Compat\Reader\IReader::SKIP_EMPTY_CELLS);
        T::same(4, \EasyExcel\Compat\Reader\IReader::IGNORE_ROWS_WITH_NO_CELLS);
    },

    // ------------------------------------------------------------------ Settings

    'wave52: setChartRenderer is accepted and round-trips' => function (): void {
        Settings::reset();
        T::same(null, Settings::getChartRenderer(), 'unset by default');
        Settings::setChartRenderer('PhpOffice\PhpSpreadsheet\Chart\Renderer\JpGraph');
        T::same(
            'PhpOffice\PhpSpreadsheet\Chart\Renderer\JpGraph',
            Settings::getChartRenderer(),
            'accepted, not thrown — charts are native (COMPAT.md)',
        );
        Settings::unsetChartRenderer();
        T::same(null, Settings::getChartRenderer(), 'unset clears it');
    },

    'wave52: setChartRenderer does not throw for an unknown class' => function (): void {
        Settings::reset();
        Settings::setChartRenderer('Totally\\Made\\Up\\Renderer');
        T::same('Totally\\Made\\Up\\Renderer', Settings::getChartRenderer(), 'no validation, no throw');
    },

    'wave52: locale and cache accessors round-trip' => function (): void {
        Settings::reset();
        // Upstream's real surface — an earlier revision invented libxml
        // accessors that PhpSpreadsheet does not have.
        T::same('en', Settings::getLocale(), 'default locale');
        T::same(true, Settings::setLocale('fr'), 'upstream returns success/failure');
        T::same('fr', Settings::getLocale());
        T::same(\ENT_COMPAT, Settings::htmlEntityFlags(), 'matches upstream');

        $cache = new \stdClass();
        Settings::setCache($cache);
        T::ok(Settings::getCache() === $cache, 'cache round-trips by identity');
    },

    'wave52: http client and request factory round-trip together' => function (): void {
        Settings::reset();
        $client = new \stdClass();
        $factory = new \stdClass();
        Settings::setHttpClient($client, $factory);
        T::ok(Settings::getHttpClient() === $client, 'client');
        T::ok(Settings::getRequestFactory() === $factory, 'factory');
        Settings::unsetHttpClient();
        T::same(null, Settings::getHttpClient(), 'unset clears client');
        T::same(null, Settings::getRequestFactory(), 'unset clears factory too');
    },

    'wave52: Settings::reset restores every default' => function (): void {
        Settings::setChartRenderer('X');
        Settings::setLocale('de');
        Settings::setCache(new \stdClass());
        Settings::reset();
        T::same(null, Settings::getChartRenderer());
        T::same('en', Settings::getLocale());
        T::same(null, Settings::getCache());
    },

    // --------------------------------------------------------------- CellAddress

    'wave52: CellAddress::fromCellAddress parses column and row' => function (): void {
        $a = CellAddress::fromCellAddress('B7');
        T::same('B', $a->columnName());
        T::same(2, $a->columnId());
        T::same(7, $a->rowId());
        T::same('B7', $a->cellAddress());
        T::same('$B$7', $a->absoluteCellAddress());
        T::same('B7', (string) $a, 'stringable');
    },

    'wave52: CellAddress handles multi-letter columns' => function (): void {
        $a = CellAddress::fromCellAddress('AA10');
        T::same('AA', $a->columnName());
        T::same(27, $a->columnId());
        T::same(10, $a->rowId());
    },

    'wave52: CellAddress::fromColumnAndRow round-trips' => function (): void {
        $a = CellAddress::fromColumnAndRow(3, 5);
        T::same('C5', $a->cellAddress());
        T::same('C', $a->columnName());
        T::same(3, $a->columnId());
    },

    'wave52: CellAddress rejects zero and negative indexes' => function (): void {
        T::throws(CompatException::class, static fn () => CellAddress::fromColumnAndRow(0, 1), 'column 0');
        T::throws(CompatException::class, static fn () => CellAddress::fromColumnAndRow(1, 0), 'row 0');
        T::throws(CompatException::class, static fn () => CellAddress::fromColumnAndRow(-1, 5), 'negative column');
    },

    'wave52: CellAddress navigation returns new instances' => function (): void {
        $a = CellAddress::fromCellAddress('C5');
        T::same('C6', $a->nextRow()->cellAddress());
        T::same('C4', $a->previousRow()->cellAddress());
        T::same('D5', $a->nextColumn()->cellAddress());
        T::same('B5', $a->previousColumn()->cellAddress());
        T::same('C7', $a->nextRow(2)->cellAddress(), 'offset honoured');
        T::same('C5', $a->cellAddress(), 'original is immutable');
    },

    'wave52: CellAddress navigation clamps at the sheet edge' => function (): void {
        $a = CellAddress::fromCellAddress('A1');
        T::same('A1', $a->previousRow()->cellAddress(), 'row clamps to 1');
        T::same('A1', $a->previousColumn()->cellAddress(), 'column clamps to A');
    },

    // -------------------------------------------------------------- AddressRange

    'wave52: AddressRange::fromCellRange parses both corners' => function (): void {
        $r = AddressRange::fromCellRange('B2:D9');
        T::same('B2', $r->from()->cellAddress());
        T::same('D9', $r->to()->cellAddress());
        T::same('B2:D9', $r->cellRange());
        T::same('$B$2:$D$9', $r->absoluteCellRange());
        T::same('B2:D9', (string) $r, 'stringable');
    },

    'wave52: AddressRange normalises a backwards range' => function (): void {
        $r = AddressRange::fromCellRange('D9:B2');
        T::same('B2', $r->from()->cellAddress(), 'top-left first');
        T::same('D9', $r->to()->cellAddress(), 'bottom-right second');
    },

    'wave52: AddressRange normalises mixed corners' => function (): void {
        // Top-right and bottom-left: neither given corner is the answer.
        $r = AddressRange::fromCellRange('D2:B9');
        T::same('B2', $r->from()->cellAddress());
        T::same('D9', $r->to()->cellAddress());
    },

    'wave52: AddressRange accepts a single-cell range' => function (): void {
        $r = AddressRange::fromCellRange('A1:A1');
        T::same('A1', $r->from()->cellAddress());
        T::same('A1', $r->to()->cellAddress());
    },

    'wave52: AddressRange rejects malformed ranges' => function (): void {
        T::throws(CompatException::class, static fn () => AddressRange::fromCellRange('A1'), 'no colon');
        T::throws(CompatException::class, static fn () => AddressRange::fromCellRange('A1:'), 'empty end');
        T::throws(CompatException::class, static fn () => AddressRange::fromCellRange(':B2'), 'empty start');
        T::throws(CompatException::class, static fn () => AddressRange::fromCellRange('A1:B2:C3'), 'three parts');
    },

    'wave52: AddressRange built from CellAddress objects' => function (): void {
        $r = new AddressRange(
            CellAddress::fromColumnAndRow(1, 1),
            CellAddress::fromColumnAndRow(3, 4),
        );
        T::same('A1:C4', $r->cellRange());
    },

    // --------------------------------------------------------------- BaseDrawing

    'wave52: Drawing and MemoryDrawing share the BaseDrawing parent' => function (): void {
        T::ok(new Drawing() instanceof BaseDrawing, 'Drawing');
        T::ok(new MemoryDrawing() instanceof BaseDrawing, 'MemoryDrawing');
    },

    'wave52: BaseDrawing accessors work through Drawing' => function (): void {
        $d = new Drawing();
        T::ok($d->setName('logo') === $d, 'fluent');
        T::same('logo', $d->getName());
        $d->setDescription('company logo');
        T::same('company logo', $d->getDescription());
        $d->setCoordinates('C3');
        T::same('C3', $d->getCoordinates());
        $d->setOffsetX(4)->setOffsetY(6)->setWidth(120)->setHeight(80);
        T::same(4, $d->getOffsetX());
        T::same(6, $d->getOffsetY());
        T::same(120, $d->getWidth());
        T::same(80, $d->getHeight());
        T::same(null, $d->getWorksheet(), 'unattached');
    },

    'wave52: BaseDrawing accessors work through MemoryDrawing' => function (): void {
        $d = new MemoryDrawing();
        $d->setName('chart')->setCoordinates('B2')->setWidth(50)->setHeight(25);
        T::same('chart', $d->getName());
        T::same('B2', $d->getCoordinates());
        T::same(50, $d->getWidth());
        T::same(25, $d->getHeight());
    },

    'wave52: drawing defaults match PhpSpreadsheet' => function (): void {
        $d = new Drawing();
        T::same('', $d->getName());
        T::same('', $d->getDescription());
        T::same('A1', $d->getCoordinates(), 'anchors at A1 by default');
        T::same(0, $d->getOffsetX());
        T::same(0, $d->getWidth());
    },

    'wave52: a BaseDrawing-typed parameter accepts both subclasses' => function (): void {
        $take = static fn (BaseDrawing $d): string => $d->getCoordinates();
        T::same('A1', $take(new Drawing()));
        T::same('A1', $take(new MemoryDrawing()));
    },

    'wave52: Drawing still guards attaching without a path' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $d = new Drawing();
        T::throws(
            CompatException::class,
            static fn () => $d->setWorksheet($s->getActiveSheet()),
            'the refactor kept the path guard',
        );
    },

    // ------------------------------------------------------------- Shared\Drawing

    'wave52: points and pixels convert at 96 DPI' => function (): void {
        T::same(0, SharedDrawing::pointsToPixels(0), 'zero stays zero');
        T::same(16, SharedDrawing::pointsToPixels(12), '12pt -> 16px');
        T::same(9.0, SharedDrawing::pixelsToPoints(12), '12px -> 9pt');
    },

    'wave52: pixels and EMU convert at 9525' => function (): void {
        T::same(9525, SharedDrawing::pixelsToEMU(1));
        T::same(95250, SharedDrawing::pixelsToEMU(10));
        T::same(10, SharedDrawing::EMUToPixels(95250));
        T::same(0, SharedDrawing::EMUToPixels(0), 'zero stays zero');
    },

    'wave52: centimetre and inch conversions to EMU' => function (): void {
        T::same(360000, SharedDrawing::centimetersToEMU(1));
        T::same(914400, SharedDrawing::inchesToEMU(1));
    },

    'wave52: degrees convert to OOXML rotation units' => function (): void {
        T::same(2700000, SharedDrawing::degreesToAngle(45));
        T::same(45, SharedDrawing::angleToDegrees(2700000));
        T::same(0, SharedDrawing::angleToDegrees(0), 'zero stays zero');
    },

    'wave52: cellDimensionToPixels uses the font character width' => function (): void {
        $font = new FontStub('Calibri', 11);
        T::same(70, SharedDrawing::cellDimensionToPixels(10, $font), '10 chars at 7px');
        T::same(0, SharedDrawing::cellDimensionToPixels(-1, $font), 'negative means unset');
        T::same(70, SharedDrawing::cellDimensionToPixels(10, null), 'null font uses the default width');
    },

    'wave52: cellDimension conversion round-trips' => function (): void {
        $font = new FontStub('Calibri', 11);
        $px = SharedDrawing::cellDimensionToPixels(12, $font);
        T::same(12.0, SharedDrawing::pixelsToCellDimension($px, $font));
    },

    // ---------------------------------------------------------------- Shared\Font

    'wave52: default row height is looked up per font and size' => function (): void {
        $calibri11 = new FontStub('Calibri', 11);
        T::same(15.0, SharedFont::getDefaultRowHeightByFont($calibri11));

        $arial10 = new FontStub('Arial', 10);
        T::same(12.75, SharedFont::getDefaultRowHeightByFont($arial10));
    },

    'wave52: default row height falls back for unlisted fonts' => function (): void {
        $exotic = new FontStub('Comic Sans MS', 10);
        $h = SharedFont::getDefaultRowHeightByFont($exotic);
        T::ok($h > 0, 'returns a usable height rather than throwing');
        T::same(13.64, $h, 'linear approximation of the same curve');
    },

    'wave52: default row height tolerates a null font' => function (): void {
        T::same(15.0, SharedFont::getDefaultRowHeightByFont(null), 'defaults to Calibri 11');
    },

    'wave52: character width scales with font size' => function (): void {
        T::same(7, SharedFont::getCharacterWidth(null), 'default');
        T::same(7, SharedFont::getCharacterWidth(new FontStub('Calibri', 11)));
        T::same(14, SharedFont::getCharacterWidth(new FontStub('Calibri', 22)), 'doubles at 22pt');
        T::ok(SharedFont::getCharacterWidth(new FontStub('Calibri', 1)) >= 1, 'never zero');
    },

    'wave52: the real Style\\Font works through the duck-typed contract' => function (): void {
        // The production path: a font obtained from a live Style, not a stub.
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $font = $s->getActiveSheet()->getStyle('A1')->getFont();
        $font->setName('Arial')->setSize(10);

        T::same(12.75, SharedFont::getDefaultRowHeightByFont($font), 'looked up from a real Style font');
        T::same(6, SharedFont::getCharacterWidth($font), '10pt scales below the 11pt default');
        T::same(60, SharedDrawing::cellDimensionToPixels(10, $font), 'and feeds the pixel conversion');
    },

    'wave52: font helpers tolerate an object without getName/getSize' => function (): void {
        $notAFont = new \stdClass();
        T::same(15.0, SharedFont::getDefaultRowHeightByFont($notAFont), 'falls back to Calibri 11');
        T::same(7, SharedFont::getCharacterWidth($notAFont), 'falls back to the default width');
    },

    'wave52: auto-size method reports the approximation' => function (): void {
        T::same(SharedFont::AUTOSIZE_METHOD_APPROX, SharedFont::getAutoSizeMethod());
        T::same(true, SharedFont::setAutoSizeMethod(SharedFont::AUTOSIZE_METHOD_APPROX), 'approx accepted');
        T::same(false, SharedFont::setAutoSizeMethod(SharedFont::AUTOSIZE_METHOD_EXACT), 'exact declined, not thrown');
    },

    // ---------------------------------------------------------------- Shared\File

    'wave52: sysGetTempDir returns a writable directory' => function (): void {
        SharedFile::reset();
        $dir = SharedFile::sysGetTempDir();
        T::ok(\is_dir($dir), 'exists');
        T::ok(\is_writable($dir), 'writable — callers build temp filenames under it');
    },

    'wave52: temporaryFilename creates a real file in the temp dir' => function (): void {
        SharedFile::reset();
        $f = SharedFile::temporaryFilename();
        T::ok(\is_file($f), 'created');
        T::ok(\str_starts_with($f, SharedFile::sysGetTempDir()), 'lives under the temp dir');
        \unlink($f);
    },

    'wave52: temporaryFilename returns distinct paths' => function (): void {
        SharedFile::reset();
        $a = SharedFile::temporaryFilename();
        $b = SharedFile::temporaryFilename();
        T::ok($a !== $b, 'unique per call');
        \unlink($a);
        \unlink($b);
    },

    'wave52: upload temp directory toggle round-trips' => function (): void {
        SharedFile::reset();
        T::same(false, SharedFile::getUseUploadTempDirectory(), 'off by default');
        SharedFile::setUseUploadTempDirectory(false);
        T::same(false, SharedFile::getUseUploadTempDirectory(), 'still off');
        SharedFile::reset();
    },

    'wave52: fileExists and realpath behave on real and missing paths' => function (): void {
        $self = __FILE__;
        T::same(true, SharedFile::fileExists($self));
        T::same(false, SharedFile::fileExists('/nonexistent/easy-excel/nope.txt'));
        T::same(\realpath($self), SharedFile::realpath($self));
        T::same('', SharedFile::realpath('/nonexistent/easy-excel/nope.txt'), 'missing resolves to empty');
        T::same(false, SharedFile::fileExists(\dirname($self)), 'a directory is not a file');
    },

    // ------------------------------------------------------------------- aliasing

    'wave52: every new class resolves through the PhpOffice alias' => function (): void {
        foreach ([
            'PhpOffice\PhpSpreadsheet\Writer\Exception',
            'PhpOffice\PhpSpreadsheet\Reader\Exception',
            'PhpOffice\PhpSpreadsheet\Calculation\Exception',
            'PhpOffice\PhpSpreadsheet\Reader\IReader',
            'PhpOffice\PhpSpreadsheet\Settings',
            'PhpOffice\PhpSpreadsheet\Cell\CellAddress',
            'PhpOffice\PhpSpreadsheet\Cell\AddressRange',
            'PhpOffice\PhpSpreadsheet\Worksheet\BaseDrawing',
            'PhpOffice\PhpSpreadsheet\Shared\Drawing',
            'PhpOffice\PhpSpreadsheet\Shared\Font',
            'PhpOffice\PhpSpreadsheet\Shared\File',
        ] as $class) {
            T::ok(
                \class_exists($class) || \interface_exists($class),
                "$class resolves via the bootstrap alias",
            );
        }
    },

    'wave52: aliased exceptions are catchable under the PhpOffice name' => function (): void {
        $writerAlias = 'PhpOffice\PhpSpreadsheet\Writer\Exception';
        $e = new \EasyExcel\Compat\Writer\Exception('aliased');
        T::ok($e instanceof $writerAlias, 'the alias and the Compat class are the same type');
    },

    'wave52: aliased CellAddress is usable through the PhpOffice name' => function (): void {
        $cls = 'PhpOffice\PhpSpreadsheet\Cell\CellAddress';
        $a = $cls::fromCellAddress('D4');
        T::same('D4', $a->cellAddress());
        T::ok($a instanceof CellAddress, 'same underlying class');
    },
];
