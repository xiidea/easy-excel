<?php

declare(strict_types=1);

namespace EasyExcel\Compat\Writer;

use EasyExcel\Compat\Cell\Coordinate;
use EasyExcel\Compat\Shared\StreamPath;
use EasyExcel\Compat\Spreadsheet;
use EasyExcel\Compat\Worksheet\PageSetup;
use EasyExcel\Compat\Worksheet\Worksheet;

/**
 * Renders a workbook to HTML, PhpSpreadsheet-compatible. This is pure PHP (the
 * Go/excelize extension has no HTML output) and works off the formatted cell
 * values exposed by the Compat layer, so it is available with or without the
 * extension.
 *
 * Faithful to the common PhpSpreadsheet surface — sheet tables with merged-cell
 * row/colspans, an optional sheet-navigation block, and the generate*() pieces
 * (header / styles / sheet data / navigation / footer). The fine-grained style,
 * inline-CSS, image-embedding and conditional-formatting knobs are accepted for
 * source compatibility but render with a single shared stylesheet (COMPAT.md).
 */
class Html extends BaseWriter
{
    protected ?int $sheetIndex = 0;

    private bool $generateSheetNavigationBlock = true;

    private bool $useInlineCss = false;

    private bool $embedImages = false;

    private string $imagesRoot = '.';

    private string $lineEnding = PHP_EOL;

    private bool $betterBoolean = true;

    private bool $tableFormats = false;

    private bool $conditionalFormatting = false;

    private bool $dataFormula = false;

    private bool $preserveFormatAndValue = false;

    /** @var null|callable(string):string */
    private $editHtmlCallback;

    protected array $cssCache = [];

    public function __construct(private Spreadsheet $spreadsheet)
    {
    }

    /**
     * @param resource|string $filename filesystem path, stream-wrapper URL, or open stream
     */
    public function save($filename, int $flags = 0): void
    {
        $this->processFlags($flags);
        $html = $this->generateHtmlAll();

        if (\is_string($filename) && !StreamPath::isWrapped($filename)) {
            if (\file_put_contents($filename, $html) === false) {
                throw new Exception("Could not write to $filename");
            }

            return;
        }

        $this->openFileHandle($filename);
        \fwrite($this->fileHandle, $html);
        $this->maybeCloseFileHandle();
    }

    /** Full standalone document: header (with styles), navigation, sheet data, footer. */
    public function generateHtmlAll(): string
    {
        $this->cssCache = [];
        $data = $this->generateSheetData();

        $html = $this->generateHTMLHeader(true);
        if ($this->generateSheetNavigationBlock && $this->sheetIndex === null) {
            $html .= $this->generateNavigation();
        }
        $html .= $data;
        $html .= $this->generateHTMLFooter();

        if ($this->editHtmlCallback !== null) {
            $html = ($this->editHtmlCallback)($html);
        }

        return $html;
    }

    public function generateHTMLHeader(bool $includeStyles = false): string
    {
        $title = $this->spreadsheet->getProperties()->getTitle();
        $eol = $this->lineEnding;

        $html = '<!DOCTYPE html>' . $eol
            . '<html lang="en">' . $eol
            . '<head>' . $eol
            . '<meta charset="UTF-8" />' . $eol
            . '<title>' . self::escape($title !== '' ? $title : 'Spreadsheet') . '</title>' . $eol;
        if ($includeStyles) {
            $html .= $this->generateStyles(false);
        }
        $html .= '</head>' . $eol
            . '<body>' . $eol;

        return $html;
    }

    public function generateStyles(bool $generateSurroundingHTML = true): string
    {
        $eol = $this->lineEnding;
        $css = 'body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; }' . $eol
            . 'table.sheet { border-collapse: collapse; margin-bottom: 1em; }' . $eol
            . 'table.sheet caption { font-weight: bold; text-align: left; padding: 0.25em 0; }' . $eol
            . 'table.sheet td { border: 1px solid #d0d0d0; padding: 1px 3px; vertical-align: top; }' . $eol
            . 'nav.sheet-navigation a { margin-right: 1em; }' . $eol
            . 'thead { display: table-header-group; }' . $eol
            . 'tfoot { display: table-footer-group; }' . $eol
            . 'tr { page-break-inside: avoid; }' . $eol
            . 'table.sheet { margin-bottom: 0; }' . $eol
            . 'table.sheet + table.sheet { border-top: none; }' . $eol;

        // Unique CSS classes cache styles
        foreach ($this->cssCache as $hash => $style) {
            $css .= 'table.sheet .' . $style['class'] . ' { ' . \implode(' ', $style['rules']) . ' }' . $eol;
        }

        if (!$generateSurroundingHTML) {
            return '<style type="text/css">' . $eol . $css . '</style>' . $eol;
        }

        return $this->generateHTMLHeader(false)
            . '<style type="text/css">' . $eol . $css . '</style>' . $eol
            . $this->generateHTMLFooter();
    }

    public function generateNavigation(): string
    {
        $eol = $this->lineEnding;
        $html = '<nav class="sheet-navigation">' . $eol;
        foreach ($this->spreadsheet->getAllSheets() as $i => $sheet) {
            $html .= '<a href="#sheet' . $i . '">' . self::escape($sheet->getTitle()) . '</a>' . $eol;
        }
        $html .= '</nav>' . $eol;

        return $html;
    }

    public function generateSheetData(): string
    {
        $html = '';
        foreach ($this->sheetsToRender() as $index => $sheet) {
            $html .= $this->generateTable($sheet, $index);
        }

        return $html;
    }

    public function generateHTMLFooter(): string
    {
        return '</body>' . $this->lineEnding . '</html>' . $this->lineEnding;
    }

    // -- accessors (PhpSpreadsheet-compatible) ---------------------------------

    public function getSheetIndex(): ?int
    {
        return $this->sheetIndex;
    }

    public function setSheetIndex(int $sheetIndex): static
    {
        $this->sheetIndex = $sheetIndex;

        return $this;
    }

    /** Render every sheet instead of a single one. */
    public function writeAllSheets(): static
    {
        $this->sheetIndex = null;

        return $this;
    }

    public function getGenerateSheetNavigationBlock(): bool
    {
        return $this->generateSheetNavigationBlock;
    }

    public function setGenerateSheetNavigationBlock(bool $generateSheetNavigationBlock): static
    {
        $this->generateSheetNavigationBlock = $generateSheetNavigationBlock;

        return $this;
    }

    public function getUseInlineCss(): bool
    {
        return $this->useInlineCss;
    }

    public function setUseInlineCss(bool $useInlineCss): static
    {
        $this->useInlineCss = $useInlineCss;

        return $this;
    }

    public function getEmbedImages(): bool
    {
        return $this->embedImages;
    }

    public function setEmbedImages(bool $embedImages): static
    {
        $this->embedImages = $embedImages;

        return $this;
    }

    public function getImagesRoot(): string
    {
        return $this->imagesRoot;
    }

    public function setImagesRoot(string $imagesRoot): static
    {
        $this->imagesRoot = $imagesRoot;

        return $this;
    }

    public function getLineEnding(): string
    {
        return $this->lineEnding;
    }

    public function setLineEnding(string $lineEnding): self
    {
        $this->lineEnding = $lineEnding;

        return $this;
    }

    public function getTableFormats(): bool
    {
        return $this->tableFormats;
    }

    public function setTableFormats(bool $tableFormats, ?bool $tableFormatsBuiltin = null): self
    {
        $this->tableFormats = $tableFormats;

        return $this;
    }

    public function getConditionalFormatting(): bool
    {
        return $this->conditionalFormatting;
    }

    public function setConditionalFormatting(bool $conditionalFormatting): self
    {
        $this->conditionalFormatting = $conditionalFormatting;

        return $this;
    }

    public function getBetterBoolean(): bool
    {
        return $this->betterBoolean;
    }

    public function setBetterBoolean(bool $betterBoolean): self
    {
        $this->betterBoolean = $betterBoolean;

        return $this;
    }

    public function setDataFormula(bool $dataFormula): self
    {
        $this->dataFormula = $dataFormula;

        return $this;
    }

    public function setPreserveFormatAndValue(bool $preserveFormatAndValue): self
    {
        $this->preserveFormatAndValue = $preserveFormatAndValue;

        return $this;
    }

    public function setEditHtmlCallback(?callable $callback): void
    {
        $this->editHtmlCallback = $callback;
    }

    /** Page orientation of the rendered sheet, or null when left at default. */
    public function getOrientation(): ?string
    {
        $sheet = $this->spreadsheet->getSheet($this->sheetIndex ?? 0);
        $orientation = $sheet->getPageSetup()->getOrientation();

        return $orientation === PageSetup::ORIENTATION_DEFAULT ? null : $orientation;
    }

    // -- internals -------------------------------------------------------------

    /** @return array<int, Worksheet> sheets to render keyed by their workbook index */
    private function sheetsToRender(): array
    {
        if ($this->sheetIndex === null) {
            return $this->spreadsheet->getAllSheets();
        }

        return [$this->sheetIndex => $this->spreadsheet->getSheet($this->sheetIndex)];
    }

    private function generateTable(Worksheet $sheet, int $index): string
    {
        $eol = $this->lineEnding;
        $rows = $sheet->toArray(null, true, true, false);
        $spans = $this->mergeSpans($sheet);
        $sheetName = $sheet->getTitle();

        $repeatRows = [];
        if ($sheet->getPageSetup()->isRowsToRepeatAtTopSet()) {
            $repeatRows = $sheet->getPageSetup()->getRowsToRepeatAtTop();
        } else {
            // Fallback: search defined names for '_xlnm.Print_Titles' if sheet was loaded from a file
            $definedNames = $sheet->getParent()->getDefinedNames();
            $printTitlesRange = $definedNames['_XLNM.PRINT_TITLES'] ?? $definedNames['_xlnm.Print_Titles'] ?? null;
            if ($printTitlesRange !== null) {
                $scopedSheet = $printTitlesRange->getWorksheet();
                if ($scopedSheet === null || $scopedSheet->getTitle() === $sheetName) {
                    $refersTo = $printTitlesRange->getRange();
                    foreach (\explode(',', $refersTo) as $part) {
                        if (\preg_match('/\$(\d+):\$(\d+)/', $part, $matches)) {
                            $repeatRows = [(int)$matches[1], (int)$matches[2]];
                            break;
                        }
                    }
                }
            }
        }

        $hasRepeat = !empty($repeatRows) && isset($repeatRows[0], $repeatRows[1]);
        $repeatStart = $hasRepeat ? (int)$repeatRows[0] : 0;
        $repeatEnd = $hasRepeat ? (int)$repeatRows[1] : 0;
        $highestRow = \count($rows);
        $highestColIndex = Coordinate::columnIndexFromString($sheet->getHighestColumn());

        // Get Column Widths
        $colWidths = [];
        for ($col = 1; $col <= $highestColIndex; $col++) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            $dim = $sheet->getColumnDimension($colLetter);
            $w = $dim->getWidth();
            if ($w > 0) {
                $colWidths[$col] = $w;
            }
        }

        // Get Row Heights
        $rowHeights = [];
        for ($row = 1; $row <= $highestRow; $row++) {
            $dim = $sheet->getRowDimension($row);
            $h = $dim->getRowHeight();
            if ($h > 0) {
                $rowHeights[$row] = $h;
            }
        }

        // Get Drawings (Images) map
        $drawings = $sheet->getDrawingCollection();
        $drawingMap = [];
        foreach ($drawings as $d) {
            $drawingMap[$d->getCoordinates()][] = $d;
        }

        $renderColGroup = function () use ($colWidths, $highestColIndex): string {
            $cg = '';
            if (!empty($colWidths)) {
                $cg .= '            <colgroup>' . PHP_EOL;
                for ($c = 1; $c <= $highestColIndex; $c++) {
                    $widthStyle = isset($colWidths[$c]) ? ' style="width: ' . \round($colWidths[$c] * 8) . 'px;"' : '';
                    $cg .= '                <col' . $widthStyle . '>' . PHP_EOL;
                }
                $cg .= '            </colgroup>' . PHP_EOL;
            }
            return $cg;
        };

        $renderRow = function (int $r) use ($rows, $spans, $eol, $sheet, $rowHeights, $drawingMap): string {
            $rowNum = $r + 1; // toArray() is a 0-indexed list; cells are 1-based
            if (!isset($rows[$r])) {
                return '';
            }
            $cells = $rows[$r];

            $rowStyle = '';
            if (isset($rowHeights[$rowNum])) {
                $rowStyle = ' style="height: ' . \round($rowHeights[$rowNum] * 1.3) . 'px;"';
            }

            $rowHtml = '<tr' . $rowStyle . '>' . $eol;
            foreach ($cells as $c => $value) {
                $colNum = $c + 1;
                $span = $spans[$rowNum][$colNum] ?? null;
                if ($span === 'covered') {
                    continue; // swallowed by a merge anchor
                }
                $attr = '';
                if (\is_array($span)) {
                    if ($span['cols'] > 1) {
                        $attr .= ' colspan="' . $span['cols'] . '"';
                    }
                    if ($span['rows'] > 1) {
                        $attr .= ' rowspan="' . $span['rows'] . '"';
                    }
                }

                $coordinate = Coordinate::stringFromColumnIndex($colNum) . $rowNum;
                $cellStyle = $sheet->getStyle($coordinate)->describe();

                $styleAttr = '';
                if ($this->useInlineCss) {
                    $rules = $this->buildCssRules($cellStyle);
                    if (!empty($rules)) {
                        $styleAttr = ' style="' . \implode(' ', $rules) . '"';
                    }
                } else {
                    $cssClass = $this->getOrCreateCssClass($cellStyle);
                    if ($cssClass !== '') {
                        $styleAttr = ' class="' . $cssClass . '"';
                    }
                }
                $attr .= $styleAttr;

                // Drawings (Images) in this cell coordinate
                $drawingsHtml = '';
                if (isset($drawingMap[$coordinate])) {
                    foreach ($drawingMap[$coordinate] as $drawing) {
                        $imgPath = $drawing->getPath();
                        if (\file_exists($imgPath)) {
                            if ($this->embedImages) {
                                $mime = \mime_content_type($imgPath) ?: 'image/png';
                                $imgData = \base64_encode(\file_get_contents($imgPath));
                                $src = "data:{$mime};base64,{$imgData}";
                            } else {
                                $src = $this->imagesRoot . '/' . \basename($imgPath);
                            }
                            $imgStyle = '';
                            if ($drawing->getName() !== '') {
                                $imgStyle .= ' alt="' . \htmlspecialchars($drawing->getName(), ENT_QUOTES, 'UTF-8') . '"';
                            }
                            $drawingsHtml .= '<img src="' . $src . '"' . $imgStyle . ' style="vertical-align: middle;">';
                        }
                    }
                }

                $cell = $value === null ? '' : self::escape((string) $value);
                $cell = \str_replace("\n", '<br>', $cell);

                if ($drawingsHtml !== '') {
                    if ($cell !== '') {
                        $cell = $drawingsHtml . '<br>' . $cell;
                    } else {
                        $cell = $drawingsHtml;
                    }
                }
                if ($cell === '') {
                    $cell = '&nbsp;';
                }

                $rowHtml .= '<td' . $attr . '>' . $cell . '</td>' . $eol;
            }
            $rowHtml .= '</tr>' . $eol;
            return $rowHtml;
        };

        $html = '';

        if ($hasRepeat) {
            // Write table 1 for rows before repeat start, if any
            if ($repeatStart > 1) {
                $html .= '<table class="sheet" id="sheet' . $index . '">' . $eol
                    . '<caption>' . self::escape($sheetName) . '</caption>' . $eol;
                $html .= $renderColGroup();
                $html .= '<tbody>' . $eol;
                for ($r = 0; $r < $repeatStart - 1; $r++) {
                    $html .= $renderRow($r);
                }
                $html .= '</tbody>' . $eol;
                $html .= '</table>' . $eol;
            }

            // A rowspan cannot cross the thead/tbody boundary: the browser
            // clamps it at the end of the section, which would drop the
            // covered cells out of alignment. Extend the header section to
            // cover any merge that starts inside it.
            //
            // Iterated to a fixed point, not a single pass: extending the
            // boundary can pull in a further merge that was previously outside
            // it. With A3:A5 and A5:A7 and repeatEnd = 3, a single pass that
            // happened to see A5:A7 first would skip it, then raise repeatEnd
            // to 5 — leaving A5:A7 straddling the boundary, which is the case
            // being fixed. Merge count is small and bounded, so re-scanning is
            // cheap; the loop is also bounded by it (review item 7).
            $merges = $sheet->getMergeCells();
            for ($pass = \count($merges); $pass > 0; --$pass) {
                $extended = false;
                foreach ($merges as $mergeRange) {
                    [[, $mergeStartRow], [, $mergeEndRow]] = Coordinate::rangeBoundaries($mergeRange);
                    if ($mergeStartRow >= $repeatStart && $mergeStartRow <= $repeatEnd && $mergeEndRow > $repeatEnd) {
                        $repeatEnd = $mergeEndRow;
                        $extended = true;
                    }
                }
                if (!$extended) {
                    break;
                }
            }

            // Write table 2 with thead (repeat rows) and tbody (subsequent rows)
            $html .= '<table class="sheet" id="sheet' . $index . '_repeated">' . $eol;
            if ($repeatStart <= 1) {
                // If it's the only table, include the caption
                $html .= '<caption>' . self::escape($sheetName) . '</caption>' . $eol;
            }
            $html .= $renderColGroup();
            $html .= '<thead>' . $eol;
            for ($r = $repeatStart - 1; $r < \min($repeatEnd, $highestRow); $r++) {
                $html .= $renderRow($r);
            }
            $html .= '</thead>' . $eol;

            if ($repeatEnd < $highestRow) {
                $html .= '<tbody>' . $eol;
                for ($r = $repeatEnd; $r < $highestRow; $r++) {
                    $html .= $renderRow($r);
                }
                $html .= '</tbody>' . $eol;
            }
            $html .= '</table>' . $eol;
        } else {
            // Write standard single table with tbody
            $html .= '<table class="sheet" id="sheet' . $index . '">' . $eol
                . '<caption>' . self::escape($sheetName) . '</caption>' . $eol;
            $html .= $renderColGroup();
            $html .= '<tbody>' . $eol;
            for ($r = 0; $r < $highestRow; $r++) {
                $html .= $renderRow($r);
            }
            $html .= '</tbody>' . $eol;
            $html .= '</table>' . $eol;
        }

        return $html;
    }

    protected function getOrCreateCssClass(array $style): string
    {
        if (empty($style)) {
            return '';
        }

        $hash = \md5(\json_encode($style));
        if (isset($this->cssCache[$hash])) {
            return $this->cssCache[$hash]['class'];
        }

        $rules = $this->buildCssRules($style);
        if (empty($rules)) {
            return '';
        }

        $className = 'style_' . \count($this->cssCache);
        $this->cssCache[$hash] = [
            'class' => $className,
            'rules' => $rules
        ];

        return $className;
    }

    protected function buildCssRules(array $style): array
    {
        $rules = [];

        // Font
        if (isset($style['font'])) {
            $font = $style['font'];
            if (isset($font['name'])) {
                $rules[] = "font-family: '" . $font['name'] . "', sans-serif;";
            }
            if (isset($font['size'])) {
                $rules[] = "font-size: " . $font['size'] . "pt;";
            }
            if (!empty($font['bold'])) {
                $rules[] = "font-weight: bold;";
            }
            if (!empty($font['italic'])) {
                $rules[] = "font-style: italic;";
            }
            if (!empty($font['underline']) && $font['underline'] !== 'none') {
                $rules[] = "text-decoration: underline;";
            }
            if (isset($font['color']['rgb'])) {
                $rules[] = "color: #" . $font['color']['rgb'] . ";";
            } elseif (isset($font['color']['argb'])) {
                $rules[] = "color: #" . \substr($font['color']['argb'], 2) . ";";
            }
        }

        // Fill / Background
        if (isset($style['fill'])) {
            $fill = $style['fill'];
            if (isset($fill['fillType']) && $fill['fillType'] === 'solid') {
                if (isset($fill['startColor']['rgb'])) {
                    $rules[] = "background-color: #" . $fill['startColor']['rgb'] . ";";
                } elseif (isset($fill['startColor']['argb'])) {
                    $rules[] = "background-color: #" . \substr($fill['startColor']['argb'], 2) . ";";
                }
            }
        }

        // Alignment
        if (isset($style['alignment'])) {
            $align = $style['alignment'];
            if (isset($align['horizontal'])) {
                $h = $align['horizontal'];
                if ($h === 'center') {
                    $rules[] = "text-align: center;";
                } elseif ($h === 'right') {
                    $rules[] = "text-align: right;";
                } elseif ($h === 'left') {
                    $rules[] = "text-align: left;";
                } elseif ($h === 'justify') {
                    $rules[] = "text-align: justify;";
                }
            }
            if (isset($align['vertical'])) {
                $v = $align['vertical'];
                if ($v === 'center') {
                    $rules[] = "vertical-align: middle;";
                } elseif ($v === 'bottom') {
                    $rules[] = "vertical-align: bottom;";
                } elseif ($v === 'top') {
                    $rules[] = "vertical-align: top;";
                }
            }
            if (!empty($align['wrapText'])) {
                $rules[] = "white-space: normal; word-wrap: break-word;";
            }
        }

        // Borders
        if (isset($style['borders'])) {
            // allBorders applies to every side unless a side overrides it.
            $allBorders = $style['borders']['allBorders'] ?? null;
            foreach (['top', 'bottom', 'left', 'right'] as $borderName) {
                $border = $style['borders'][$borderName] ?? $allBorders;
                if (null !== $border) {
                    $borderStyle = $border['borderStyle'] ?? 'none';
                    if ($borderStyle === 'none') {
                        // An explicit BORDER_NONE must emit a rule: the
                        // sheet-wide `table.sheet td` border would otherwise
                        // stay visible, since nothing overrides it. Cells that
                        // were never styled have no 'borders' key at all, so
                        // they keep the sheet default.
                        $rules[] = "border-{$borderName}: none;";
                    } else {
                        $width = '1px';
                        if (\str_contains($borderStyle, 'medium')) {
                            $width = '2px';
                        } elseif (\str_contains($borderStyle, 'thick')) {
                            $width = '3px';
                        }

                        $type = 'solid';
                        if (\str_contains($borderStyle, 'dashed')) {
                            $type = 'dashed';
                        } elseif (\str_contains($borderStyle, 'dotted')) {
                            $type = 'dotted';
                        } elseif (\str_contains($borderStyle, 'double')) {
                            $type = 'double';
                        }

                        $color = 'black';
                        if (isset($border['color']['rgb'])) {
                            $color = "#" . $border['color']['rgb'];
                        } elseif (isset($border['color']['argb'])) {
                            $color = "#" . \substr($border['color']['argb'], 2);
                        }
                        $rules[] = "border-{$borderName}: {$width} {$type} {$color};";
                    }
                }
            }
        }

        return $rules;
    }

    /**
     * Resolve merged ranges into per-cell span info: the top-left anchor carries
     * {rows, cols}; every other covered cell is marked 'covered' so it is skipped.
     *
     * @return array<int, array<int, array{rows:int, cols:int}|string>>
     */
    private function mergeSpans(Worksheet $sheet): array
    {
        $spans = [];
        foreach ($sheet->getMergeCells() as $range) {
            [[$startCol, $startRow], [$endCol, $endRow]] = Coordinate::rangeBoundaries($range);
            $spans[$startRow][$startCol] = ['rows' => $endRow - $startRow + 1, 'cols' => $endCol - $startCol + 1];
            for ($row = $startRow; $row <= $endRow; ++$row) {
                $row = (int) $row;
                for ($col = $startCol; $col <= $endCol; ++$col) {
                    $col = (int) $col;
                    if ($row === $startRow && $col === $startCol) {
                        continue;
                    }
                    $spans[$row][$col] = 'covered';
                }
            }
        }

        return $spans;
    }

    private static function escape(string $value): string
    {
        return \htmlspecialchars($value, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
    }
}
