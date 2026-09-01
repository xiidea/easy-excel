<?php

declare(strict_types=1);

use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Shared\StringHelper;
use EasyExcel\Compat\Writer\BaseWriter;

/**
 * Stands in for the consumer apps' HTMLWriter after wave 5.1: a standalone
 * renderer that extends BaseWriter directly and declares everything it uses,
 * rather than subclassing Writer\Html for internals it never inherited.
 *
 * This is the contract wave 5.1 relies on — if BaseWriter stops carrying the
 * file-handle helpers or the flag properties, both apps break and this fails.
 */
final class StandaloneWriter extends BaseWriter
{
    public function __construct(private Spreadsheet $spreadsheet)
    {
    }

    public function generateBody(): string
    {
        $ws = $this->spreadsheet->getActiveSheet();
        $rows = '';
        foreach ($ws->toArray() as $row) {
            $rows .= '<tr><td>' . \implode('</td><td>', \array_map(
                static fn ($v) => \htmlspecialchars((string) $v),
                $row,
            )) . '</td></tr>';
        }

        return '<table>' . $rows . '</table>';
    }

    public function save($filename, int $flags = 0): void
    {
        $this->processFlags($flags);
        $this->openFileHandle($filename);
        \fwrite($this->fileHandle, $this->generateBody());
        $this->maybeCloseFileHandle();
    }
}

return [
    'wave51: a standalone BaseWriter subclass is concrete and constructible' => function (): void {
        EasyExcelFake::reset();
        $w = new StandaloneWriter(new Spreadsheet());
        T::ok($w instanceof BaseWriter, 'extends BaseWriter');
        T::ok($w instanceof \EasyExcel\Compat\Writer\IWriter, 'satisfies the IWriter contract');
    },

    'wave51: BaseWriter supplies the file-handle helpers the writers call' => function (): void {
        $rc = new \ReflectionClass(BaseWriter::class);
        foreach (['processFlags', 'openFileHandle', 'maybeCloseFileHandle'] as $m) {
            T::ok($rc->hasMethod($m), "BaseWriter::$m exists");
        }
        T::ok($rc->hasProperty('fileHandle'), 'BaseWriter::$fileHandle exists');
    },

    'wave51: BaseWriter supplies the flag properties the writers read' => function (): void {
        $rc = new \ReflectionClass(BaseWriter::class);
        T::ok($rc->hasProperty('includeCharts'), '$includeCharts');
        T::ok($rc->hasProperty('preCalculateFormulas'), '$preCalculateFormulas');
    },

    'wave51: BaseWriter has no constructor, so subclasses need no parent call' => function (): void {
        $rc = new \ReflectionClass(BaseWriter::class);
        T::same(null, $rc->getConstructor(), 'a parent::__construct() call would fatal');
    },

    'wave51: the standalone writer saves real content through BaseWriter' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $ws = $s->getActiveSheet();
        $ws->fromArray([['alpha', 1], ['beta', 2]], null, 'A1', true);
        $ws->flush();

        $file = \tempnam(\sys_get_temp_dir(), 'wave51');
        (new StandaloneWriter($s))->save($file);

        $html = (string) \file_get_contents($file);
        \unlink($file);

        T::ok(\str_contains($html, '<table>'), 'wrote through the inherited file handle');
        T::ok(\str_contains($html, 'alpha'), 'row data present');
        T::ok(\str_contains($html, 'beta'), 'second row present');
    },

    'wave51: the standalone writer escapes cell content' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setCellValue('A1', '<script>x</script>');
        $s->getActiveSheet()->flush();

        $html = (new StandaloneWriter($s))->generateBody();
        T::ok(!\str_contains($html, '<script>'), 'raw tag not emitted');
        T::ok(\str_contains($html, '&lt;script&gt;'), 'escaped instead');
    },

    'wave51: BaseWriter flag accessors round-trip on the subclass' => function (): void {
        EasyExcelFake::reset();
        $w = new StandaloneWriter(new Spreadsheet());
        T::same(false, $w->getIncludeCharts(), 'default');
        $w->setIncludeCharts(true);
        T::same(true, $w->getIncludeCharts());

        T::same(true, $w->getPreCalculateFormulas(), 'default');
        $w->setPreCalculateFormulas(false);
        T::same(false, $w->getPreCalculateFormulas());
    },

    'wave51: save() accepts an open stream as well as a path' => function (): void {
        EasyExcelFake::reset();
        $s = new Spreadsheet();
        $s->getActiveSheet()->setCellValue('A1', 'streamed');
        $s->getActiveSheet()->flush();

        $file = \tempnam(\sys_get_temp_dir(), 'wave51s');
        $fh = \fopen($file, 'wb');
        (new StandaloneWriter($s))->save($fh);
        // The writer must not close a handle it did not open.
        T::ok(\is_resource($fh), 'caller-owned handle stays open');
        \fclose($fh);

        T::ok(\str_contains((string) \file_get_contents($file), 'streamed'), 'content written');
        \unlink($file);
    },

    // ------------------------------------------------- Shared\StringHelper
    // Verified byte-identical to real PhpSpreadsheet for every case below.

    'wave51: stringIncrement walks column letters without deprecation' => function (): void {
        $cases = ['A' => 'B', 'Z' => 'AA', 'AZ' => 'BA', 'ZZ' => 'AAA', 'C' => 'D'];
        foreach ($cases as $in => $want) {
            $s = (string) $in;
            $ret = StringHelper::stringIncrement($s);
            T::same($want, $s, "$in increments in place");
            T::same($want, $ret, "$in is also returned");
        }
    },

    'wave51: stringIncrement seeds an empty string' => function (): void {
        $s = '';
        T::same('A', StringHelper::stringIncrement($s), 'empty starts at A');
    },

    'wave51: formatNumber matches upstream, including empty for null' => function (): void {
        T::same('0.7', StringHelper::formatNumber(0.7));
        T::same('1', StringHelper::formatNumber(1.0), 'whole floats lose the .0');
        T::same('0.75', StringHelper::formatNumber(0.75));
        T::same('0', StringHelper::formatNumber(0));
        T::same('12.3456789', StringHelper::formatNumber(12.3456789));
        T::same('', StringHelper::formatNumber(null), 'upstream returns empty, not 0');
        T::same('', StringHelper::formatNumber(''));
        T::same('abc', StringHelper::formatNumber('abc'), 'non-numeric passes through');
    },

    'wave51: control characters escape to the OOXML form and back' => function (): void {
        T::same('plain', StringHelper::controlCharacterPHP2OOXML('plain'), 'clean text untouched');
        T::same('_x0000_', StringHelper::controlCharacterPHP2OOXML("\x00"));
        T::same('_x0007_', StringHelper::controlCharacterPHP2OOXML("\x07"));
        T::same("a\x01b", StringHelper::controlCharacterOOXML2PHP('a_x0001_b'), 'reverses');
        $orig = "a\x01b\x1fc";
        T::same($orig, StringHelper::controlCharacterOOXML2PHP(
            StringHelper::controlCharacterPHP2OOXML($orig),
        ), 'round-trips');
    },

    'wave51: tab and newline survive control-character escaping' => function (): void {
        // \t \n \r are legal in OOXML text and must not be escaped.
        T::same("a\tb", StringHelper::controlCharacterPHP2OOXML("a\tb"));
        T::same("a\nb", StringHelper::controlCharacterPHP2OOXML("a\nb"));
        T::same("a\rb", StringHelper::controlCharacterPHP2OOXML("a\rb"));
    },

    'wave51: convertToString handles scalars, null and Stringable' => function (): void {
        T::same('123', StringHelper::convertToString(123));
        T::same('abc', StringHelper::convertToString('abc'));
        T::same('1.5', StringHelper::convertToString(1.5));
        T::same('1', StringHelper::convertToString(true));
        T::same('', StringHelper::convertToString(null));
        $obj = new class implements \Stringable {
            public function __toString(): string
            {
                return 'from-object';
            }
        };
        T::same('from-object', StringHelper::convertToString($obj));
    },

    'wave51: convertToString honours the bool and throw flags' => function (): void {
        T::same('TRUE', StringHelper::convertToString(true, true, '', true));
        T::same('FALSE', StringHelper::convertToString(false, true, '', true));
        T::same('fallback', StringHelper::convertToString([1, 2], false, 'fallback'), 'default on non-stringable');
        T::throws(
            \TypeError::class,
            static fn () => StringHelper::convertToString([1, 2]),
            'throws by default',
        );
    },

    'wave51: multibyte helpers are UTF-8 correct' => function (): void {
        T::same('HÉLLO', StringHelper::strToUpper('héllo'));
        T::same('héllo', StringHelper::strToLower('HÉLLO'));
        T::same(5, StringHelper::countCharacters('héllo'), 'characters, not bytes');
        T::same('éll', StringHelper::substring('héllo', 1, 3));
        T::same(true, StringHelper::isUTF8('héllo'));
        T::same(true, StringHelper::isUTF8(''), 'empty is valid');
    },

    'wave51: separator accessors round-trip' => function (): void {
        StringHelper::reset();
        StringHelper::setDecimalSeparator(',');
        T::same(',', StringHelper::getDecimalSeparator());
        StringHelper::setThousandsSeparator('.');
        T::same('.', StringHelper::getThousandsSeparator());
        StringHelper::reset();
        T::ok(StringHelper::getDecimalSeparator() !== '', 'reset falls back to the locale');
    },

    'wave51: the PhpOffice BaseWriter alias is what consumers extend' => function (): void {
        $alias = 'PhpOffice\PhpSpreadsheet\Writer\BaseWriter';
        T::ok(\class_exists($alias), 'alias resolves');
        T::ok(
            (new \ReflectionClass($alias))->getName() === BaseWriter::class,
            'the alias and the Compat class are the same class',
        );
    },
];
