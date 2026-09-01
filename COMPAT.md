# PhpSpreadsheet compatibility matrix

Compatibility is **measured, not asserted** (PLAN.md §5): this file tracks
what the shim implements, what intentionally diverges, and what throws a
clear "not yet supported" exception. Phase numbers refer to PLAN.md §13.

## Supported (Phase 1)

| Area | API | Notes |
|---|---|---|
| Workbook | `new Spreadsheet()`, `getActiveSheet`, `get/setActiveSheetIndex(+ByName)`, `createSheet`, `getSheet(+ByName)`, `getSheetCount/Names`, `getAllSheets`, `getIndex`, `removeSheetByIndex`, `disconnectWorksheets`, `garbageCollect` | default sheet is named `Worksheet`, like PhpSpreadsheet |
| Worksheet | `setCellValue(+ByColumnAndRow)`, `setCellValueExplicit(+ByColumnAndRow)`, `getCell(+ByColumnAndRow)`, `fromArray`, `toArray`, `rangeToArray`, `getHighestRow/Column(+Data)`, `getTitle/setTitle`, `mergeCells` | per-cell writes are buffered and batched (512 rows/CGO call) |
| Cell | `getValue`, `getCalculatedValue`, `getFormattedValue`, `setValue`, `setValueExplicit`, `getCoordinate`, `getWorksheet`, `getDataType` | data lives in Go; Cell is a coordinate facade |
| Coordinate | `columnIndexFromString`, `stringFromColumnIndex`, `coordinateFromString`, `indexesFromString`, `rangeBoundaries`, `rangeDimension`, `splitRange` | pure PHP port |
| DataType | all `TYPE_*` constants | |
| Shared\Date | `PHPToExcel`, `dateTimeToExcel`, `timestampToExcel`, `stringToExcel`, `excelToDateTimeObject`, `excelToTimestamp`, `formattedPHPToExcel`, 1900/1904 calendars | Julian-day algorithm ported verbatim, incl. the 1900 leap-year bug |
| IOFactory | `createWriter/Reader` (Xlsx, Csv, Html), `load`, `identify` | |
| Writer\IWriter, Writer\BaseWriter | full PhpSpreadsheet contract (`SAVE_WITH_CHARTS`/`DISABLE_PRECALCULATE_FORMULAE`, include-charts / pre-calculate / disk-caching accessors, `openFileHandle`/`processFlags`/`maybeCloseFileHandle`) | extend `BaseWriter` (or implement `IWriter`) for custom writers; the built-in writers extend it. Chart and disk-cache flags are state-only; **pre-calculate is wired through** (divergence 24) |
| Writer\Xlsx | `save` (paths, stream-wrapper URLs — `php://`, `gaufrette://`, `s3://`, … — and open resources) | wrapper targets are staged through a local temp file (the extension only writes real paths) |
| Writer\Csv | `set/getDelimiter`, `setEnclosure` (only `"`), `set/getLineEnding`, `set/getUseBOM`, `set/getSheetIndex`, `save` (paths, stream-wrapper URLs, open resources) | plus `setSanitizeFormulas()` (easy-excel extra, opt-in OWASP guard) |
| Writer\Html | `save`, `generateHtmlAll`, `generateHTMLHeader`, `generateStyles`, `generateNavigation`, `generateSheetData`, `generateHTMLFooter`, `set/getSheetIndex`, `writeAllSheets`, `set/getGenerateSheetNavigationBlock`, `set/getUseInlineCss`, `set/getEmbedImages`, `set/getImagesRoot`, `set/getLineEnding`, `getOrientation`, `setEditHtmlCallback`, plus the table/conditional/boolean knobs | **pure PHP** (works with or without the extension); renders formatted cell values into sheet tables with merged-cell row/colspans. Fine-grained per-cell styling and image embedding are not rendered — a single shared stylesheet is emitted |
| Reader\Xlsx | `load` (paths and stream-wrapper URLs), `setReadDataOnly`, `canRead` | wrapper sources are staged through a local temp file before the native open |
| Reader\Csv | `load`, `setDelimiter`, `setEnclosure`, `setSheetIndex`, `canRead` | streams in 1k-row chunks |
| Value binding | DefaultValueBinder semantics: numeric strings → numbers (leading-zero strings preserved), `=…` → formula, `DateTimeInterface` → Excel serial | |

## Supported (Phase 2 — formatting & structure)

| Area | API | Notes |
|---|---|---|
| Style | `getStyle(range)->applyFromArray`, `getFont` (bold/italic/size/name/underline/strikethrough/color/super/subscript), `getFill` (pattern fills + start/end color), `getBorders` (top/bottom/left/right/allBorders/outline, style + color), `getAlignment` (horizontal/vertical/wrapText/shrinkToFit/textRotation/indent), `getProtection` (locked/hidden), `getNumberFormat` | partial styles layer in application order like PhpSpreadsheet's supervisor; styles applied **before** their rows are written ride the StreamWriter at zero cost |
| Style helpers | `Color` (ARGB/RGB + constants), all `Border::BORDER_*`, `Fill::FILL_*`, `Alignment::HORIZONTAL_*/VERTICAL_*`, `Protection::PROTECTION_*`, `NumberFormat::FORMAT_*` constants | |
| Dimensions | `getColumnDimension(+ByColumn)->setWidth/setAutoSize`, `getRowDimension->setRowHeight`, `getDefaultRowDimension->setRowHeight`, `getDefaultColumnDimension->setWidth` | auto-size is approximated at save time (divergence 10); the default dimensions map to `sheetFormatPr` and never degrade a streaming sheet — prefer them over per-row heights for uniform sizing |
| Structure | `mergeCells`, `setAutoFilter`, `freezePane(+ByColumnAndRow)`, `unfreezePane` | merges/widths/panes set before streaming use the StreamWriter's native support |
| Hyperlinks | `Cell::getHyperlink()->setUrl/setTooltip`, `Worksheet::setHyperlink` | `sheet://` URLs become internal links |
| Comments | `getComment(+ByColumnAndRow)`, `Comment::setAuthor`, `getText()->createText/createTextRun/getPlainText` | plain text only; run formatting throws |
| Defined names | `Spreadsheet::addNamedRange/addDefinedName`, `NamedRange` | |
| Page setup | `getPageSetup()->setOrientation/setPaperSize/setFitToWidth/setFitToHeight/setFitToPage` | applied at save |

## Supported (Phase 3 — advanced)

| Area | API | Notes |
|---|---|---|
| Formulas | `getCalculatedValue()`, `toArray(calculateFormulas: true)`, `rangeToArray(...)` | delegated to excelize's engine: **466 of PhpSpreadsheet's 529 functions** available, per-function table in [FORMULAS.md](FORMULAS.md); cached results in loaded files are trusted |
| Data validation | `Cell::getDataValidation()` (bound, setters apply), `Cell::setDataValidation`, `Worksheet::setDataValidation(range, dv)`, all `TYPE_*`/`OPERATOR_*`/`STYLE_*` constants | list (literal + range source), whole/decimal/date/time/textLength/custom |
| Conditional formatting | `getStyle(range)->setConditionalStyles([Conditional…])`, `Conditional` (cellIs/containsText/expression + operators, `getStyle()` detached collector, `setStopIfTrue`) | plus easy-excel extras `setColorScale(min, max, mid?)` and `setDataBar(color)` (PhpSpreadsheet models these as separate classes) |
| Images | `Worksheet\Drawing`: `setName/setDescription/setPath/setCoordinates/setOffsetX/Y/setWidth/setHeight/setWorksheet` | width/height scale from the decoded PNG/JPEG/GIF dimensions; aspect kept when only one side is set |
| Sheet protection | `getProtection()->setSheet/setPassword` + all action-lock flags | applied at save; workbook encryption is not supported |
| Charts | **native API only**: `Worksheet::addNativeChart($cell, $spec)` / `Native::addChart` with a declarative spec (type, series, title, legend, size); types: area/bar/barStacked/col/colStacked/doughnut/line/pie/radar/scatter | PhpSpreadsheet's `Chart` object model is **not** mapped — see "Not yet supported" |
| Auto-filter | `setAutoFilter` on streamed sheets | now injected into the saved container (no degrade); see divergence 16 |

## Supported (Phase 4.1 — compat completion, wave 1)

| Area | API | Notes |
|---|---|---|
| Value binders | `Cell::setValueBinder/getValueBinder`, `IValueBinder`, `DefaultValueBinder` (+`dataTypeForValue`) | custom binders run in PHP before the write buffer; `fromArray` routes per cell through them (still batched); without a custom binder the bulk fast path is unchanged |
| Document properties | `getProperties()->setTitle/setSubject/setCreator/setLastModifiedBy/setDescription/setKeywords/setCategory/setCompany/setCreated/setModified`; custom properties (`setCustomProperty/getCustomProperty*/isCustomPropertySet/removeCustomProperty`, `PROPERTY_TYPE_*` constants) | `setManager` accepted but PHP-side only (excelize exposes no field); custom props persist via the docProps/custom.xml part |
| Print layout | `setRowsToRepeatAtTop(+ByStartAndEnd)`, `setColumnsToRepeatAtLeftByStartAndEnd`, `setPrintArea` | implemented as the reserved `_xlnm.Print_Titles` / `_xlnm.Print_Area` defined names |
| Conditional getter | `getStyle(range)->getConditionalStyles()` | returns rules set on that exact range **this session**; loaded files are not introspected |
| **Workbook encryption** | `Writer\Xlsx::setPassword()`, `Reader\Xlsx::setPassword()` (easy-excel extras — PhpSpreadsheet cannot encrypt xlsx) | agile encryption via excelize; encrypting disables the auto-filter container patch (filters ride the degrade) |
| Fills & borders | gradient fills (`linear`/`path` + `setRotation`), diagonal borders (`getDiagonal`, `setDiagonalDirection`, `Borders::DIAGONAL_*`) | gradient angles bucket to excelize shading directions (divergence 20) |
| Merges | `unmergeCells`, `getMergeCells()` | reading merges degrades a streaming sheet, like other reads |
| Calculation | `Calculation::getInstance()` cache controls | accepted no-ops: perf hints that cannot change output |

## Supported (Phase 4.2 — reading & iteration)

| Area | API | Notes |
|---|---|---|
| Iterators | `getRowIterator`, `getColumnIterator`, `Row::getCellIterator`, `Column::getCellIterator` (+`Row`/`Column`/`RowCellIterator`/`ColumnCellIterator`) | cells are coordinate facades reading per cell; `toArray`/`rangeToArray` remain the bulk fast path |
| Read filters | `Reader\IReadFilter`, `Reader\Xlsx::setReadFilter` | applied during chunk assembly: filtered cells read as null (PhpSpreadsheet never loads them — observable difference only via memory, which is constant here anyway) |
| Style read-back | `getStyle()` getters reflect applied styles **and loaded files** (font, fill type, alignment, number format); `Worksheet::duplicateStyle` | streaming sheets answer from the style log — read-back never degrades a workbook mid-write; loaded files reverse-translate the stylesheet |
| Default style | `Spreadsheet::getDefaultStyle()` | layered under every style fold; untouched cells get it via a full-width column style (streams through the StreamWriter) |
| Introspection | `Cell::getDataValidation()` hydrates covering rules, `getConditionalStyles()` falls back to the file's rules, `Spreadsheet::getDefinedNames()`, `Worksheet::getAutoFilter()` (session range) | validations/conditionals on streaming sheets are answered from the pending queue |

## Supported (Phase 4.3 — structure editing)

| Area | API | Notes |
|---|---|---|
| Rows/columns | `insertNewRowBefore`, `removeRow`, `insertNewColumnBefore(+ByIndex)`, `removeColumn(+ByIndex)` | random-access ops: a streaming workbook degrades first (queued styles replay before the shift, so coordinates stay valid); excelize adjusts formulas and refs |
| Sheets | `createSheet($index)` at arbitrary positions; `Spreadsheet::copySheet($source, $new)` (easy-excel extra — PhpSpreadsheet's `clone` idiom is not supported) | copy duplicates values, styles and structure |
| Sheet views | `setShowGridlines`, `getSheetView()->setZoomScale/setRightToLeft`, `getTabColor()` | applied at save |
| Print | `getHeaderFooter()` (odd/even/first headers+footers, different-first/odd-even; `&P`/`&N`/`&D`… codes pass through), `getPageMargins()` | applied at save |

## Supported (Phase 4.4 — content types)

| Area | API | Notes |
|---|---|---|
| Rich text cells | `new RichText`, `createText/createTextRun`, `Run::getFont()` (bold/italic/size/name/underline/color…), `setCellValue($coord, $richText)` | a plain placeholder keeps dimensions correct; the formatted runs apply at save (divergence 22) |
| Memory drawings | `Worksheet\MemoryDrawing` (GD resource → PNG/JPEG/GIF, `setImageResource`, `setRenderingFunction`, size/offset, `setWorksheet`) | rendered in PHP, sent to the extension as base64 bytes; requires ext-gd |
| Charts | the PhpSpreadsheet `Chart\*` object model: `Chart`, `DataSeries` (bar/column ±stacked, line, area, pie, doughnut, scatter, radar; bar/col direction), `DataSeriesValues`, `PlotArea`, `Legend`, `Title`, X/Y axis labels; `Worksheet::addChart` | mapped onto the native chart spec; series data sources are excelize formula strings |
| Auto-filter rules | `getAutoFilter()->getColumn($col)->createRule()->setRule($op, $value)`, AND/OR join | column rules force the model path (FilterColumn XML); excelize doesn't hide rows automatically (divergence 23) |

## Supported (Phase 5.2 — consumer-driven surface)

| Area | API | Notes |
|---|---|---|
| Exceptions | `Writer\Exception`, `Reader\Exception`, `Calculation\Exception` | narrow subclasses of the flat `Compat\Exception`, so `catch (Writer\Exception)` narrows correctly **and** existing broad catches keep working; the writers and readers now throw the narrow types |
| Reader contract | `Reader\IReader` (+ `READ_DATA_ONLY`/`SKIP_EMPTY_CELLS`/`IGNORE_ROWS_WITH_NO_CELLS`) | implemented by `Reader\Xlsx` and `Reader\Csv`; usable as a type hint |
| Settings | `Settings::setChartRenderer/getChartRenderer/unsetChartRenderer`, libxml + cache + HTTP-client accessors | **state-only**: values round-trip so consuming code behaves, but nothing reads them back. `setChartRenderer` is deliberately accepted rather than thrown — its callers guard an HTML/PDF preview path, and throwing would break otherwise-supported workbook generation |
| Cell addressing | `Cell\CellAddress` (`fromCellAddress`/`fromColumnAndRow`, `columnName`/`columnId`/`rowId`, `cellAddress`/`absoluteCellAddress`, `next`/`previous` row+column), `Cell\AddressRange` (`fromCellRange`, `from`/`to`, `cellRange`/`absoluteCellRange`) | immutable value objects over `Coordinate`; ranges normalise so `from()` is always top-left (`D9:B2` → `B2:D9`); navigation clamps at row 1 / column A |
| Drawings | `Worksheet\BaseDrawing` | extracted as the genuine shared parent of `Drawing` and `MemoryDrawing` (name, description, coordinates, offsets, size, owning sheet) rather than added alongside them; attachment stays per-subclass because each sends a different payload |
| Shared helpers | `Shared\StringHelper` (`stringIncrement`, `formatNumber`, `convertToString`, control-character escaping both ways, multibyte case/substring/count, separators), `Shared\Drawing` (points/pixels/EMU/cm/inch/degree conversions, `cellDimensionToPixels`, `pixelsToCellDimension`), `Shared\Font` (`getDefaultRowHeightByFont`, `getCharacterWidth`, auto-size method), `Shared\File` (`sysGetTempDir`, `temporaryFilename`, upload-temp-dir toggle, `fileExists`, `realpath`) | pure PHP. `Shared\File::sysGetTempDir()` is canonicalised with `realpath()` so it matches the paths `tempnam()` actually returns (macOS reports `/var/…` but creates under `/private/var/…`). `StringHelper` is verified byte-identical to real PhpSpreadsheet across every method the writers call, including `formatNumber(null) === ''` and `stringIncrement` via `str_increment()` (bare `++` on a string is deprecated in PHP 8.3+) |

## Supported (Phase 5.3 — page breaks & selection)

| Area | API | Notes |
|---|---|---|
| Page breaks | `Worksheet::setBreak(+ByColumnAndRow)`, `BREAK_NONE`/`BREAK_ROW`/`BREAK_COLUMN`/`BREAK_ROW_MAX_COLUMN` | `excelize.InsertPageBreak`/`RemovePageBreak`, applied at save like the rest of the non-streamable surface (divergence 11). The reference is normalised to the requested axis: `setBreak('O24', BREAK_ROW)` splits above row 24 only, never also left of column O |
| Selection | `Worksheet::setSelectedCells(+setSelectedCell/ByColumnAndRow)` | excelize carries selection inside the **pane** record, not `ViewOptions`, so the existing pane state is read back and merged — selecting a cell after `freezePane()` keeps the freeze instead of silently undoing it |
| Auto-size | `Worksheet::calculateColumnWidths()` | accepted no-op returning `$this`: auto-size is approximated in Go at save (divergence 10), so there is nothing to precompute. Callers need not branch on the engine |

## Supported (Phase 5.4 — chart axis model)

| Area | API | Notes |
|---|---|---|
| Axis | `Chart\Axis`: `setAxisOptionsProperties` (full positional signature), `setAxisNumberProperties`, `setFillParameters`, `set/getMajor+MinorGridlines`, `getAxisOptionsProperty`, `AXIS_LABELS_*`/`TICK_MARK_*`/`AXIS_ORIENTATION_*` constants | mapped onto `excelize.ChartAxis`: label suppression (`none` → `None`), `minimum`/`maximum`, `majorUnit`, `logBase`, `maxMin` → `ReverseOrder`, number format, label font colour. Bounds use pointers end-to-end so an explicit `0` is distinct from unset |
| GridLines | `Chart\GridLines`, attached via the `Chart` constructor or `Axis::setMajorGridlines` | presence turns the gridlines on (`MajorGridLines`/`MinorGridLines`). Line colour/style, glow, shadow and soft-edge setters are accepted and round-trip, but excelize models no gridline line format |
| Layout | `Chart\Layout`: `setShowVal` and the sibling data-label toggles, plot-area geometry accessors | `setShowVal` drives `PlotArea.ShowVal`; geometry is stored and round-trips but is not rendered |
| ChartColor | `Chart\ChartColor`: `setColorProperties`, `EXCEL_COLOR_TYPE_*` | normalises `#rrggbb`/`rrggbb` to bare upper-case hex for excelize |
| Chart | `getChartAxisX/Y`, `getPlotArea`, `getTitle`, `getLegend`, `getTopLeftPosition`, `setBottomRightPosition`, `render()` | gridlines passed to the constructor attach to the **Y** axis, matching PhpSpreadsheet regardless of which axis object was supplied. `setBottomRightPosition` derives an approximate pixel size (64px/column, 20px/row) since excelize sizes charts by width/height, not a second anchor. `render()` returns `false` — the value PhpSpreadsheet gives when no renderer is configured — so callers take their existing no-image branch |
| DataSeries | `EMPTY_AS_GAP`/`EMPTY_AS_ZERO`/`EMPTY_AS_SPAN`/`DEFAULT_EMPTY_AS` | accepted for constructor parity; excelize has no display-blanks-as control |

## Supported (Phase 5 cross-cutting — formula cache)

| Area | API | Notes |
|---|---|---|
| Pre-calculated results | `Writer\Xlsx::setPreCalculateFormulas(bool)`, `DISABLE_PRECALCULATE_FORMULAE` save flag | **opt-in**: the getter reports upstream's `true` default, but the pass runs only on an explicit call, so inheriting the default never costs streaming. Numeric results only; the `t="str"` excelize forces on formula cells is corrected by a container patch so number formats still apply — see divergence 24 |
| Recalculate-on-open | `Native::setFullCalcOnLoad(int $handle, bool)` (easy-excel extra, default **on**) | sets `calcPr/@fullCalcOnLoad`; costs one attribute and never degrades. Fixes spreadsheet applications; does nothing for readers that never calculate, which is what the cache above is for |

## Documented divergences

1. **`toArray(formatData: false)` types** — values come back from excelize as
   strings and are cast when the string is the *canonical* rendering of the
   number, so `"1542"` becomes `1542` while `"0042"`, `"1.50"` and `" 3"` stay
   strings. `getCalculatedValue()` follows the same rule, so a
   `TEXT(A1,"0000")` result keeps its leading zeros. Bulk `toArray` reads still
   use the looser `is_numeric()` cast, so a text cell that looks numeric can
   come back as a number there.
2. **`toArray($calculateFormulas)`** — bulk reads return raw or formatted
   values; the flag is currently honored only by `Cell::getCalculatedValue()`
   (excelize's ~535-function engine). Bulk calculated reads land in Phase 3.
3. **Formula engine coverage** — `getCalculatedValue()` delegates to excelize;
   its function set (~535) differs from PhpSpreadsheet's. A per-function table
   will be published with Phase 3.
4. **Streaming degrade** — out-of-order writes or reads on a sheet with
   already-streamed rows trigger a one-time serialize-and-reopen of the
   workbook (correct, but O(file size)). Sequential writers never hit it;
   styling no longer triggers it immediately (see divergence 9).
5. **CSV enclosure** — only `"` is supported; `setEnclosure` with anything
   else throws.
6. **`createSheet($index)`** — only appending (index = count or null);
   arbitrary insert positions throw.
7. **`setReadDataOnly`** — accepted for API parity; reads are already
   values-only fast paths.
8. **Number-format rendering** — formatted values are rendered by excelize;
   rare locale-specific format codes may render differently from
   PhpSpreadsheet. Differences found by the test suite get fixed or listed here.
9. **Style application order vs. streaming** — styles, number formats, widths,
   panes and merges applied *before* their rows are written stream at full
   speed. Styling rows that were already written queues the work and triggers
   the one-time degrade **at save** (not immediately). For big exports, style
   headers/columns first, then bulk-write.
10. **Auto-size width** — PhpSpreadsheet measures rendered text with font
    metrics; easy-excel approximates with `max character count + 2`,
    applied at save. Visually close, not byte-identical.
11. **Hyperlinks, comments, auto-size, page setup, validations, conditional
    formats, images, protection, charts** — excelize cannot stream these, so
    they are applied at save (degrading once if the sheet streamed). The data
    path itself stays streaming.
12. **Range style layering** — `getStyle('A1:C10')` applies one merged style
    per region. Earlier styles fully containing the range are folded in, and
    intersections with partially-overlapping earlier styles are re-applied
    with their own fold (so a broad late style does not clobber narrower
    earlier ones — the column-formats-then-sheet-alignment pattern works).
    Only deeper overlap chains (three-way partial overlaps whose pairwise
    intersections again partially overlap) can differ from PhpSpreadsheet's
    strict per-cell layering.
13. **Full-column styles** (`getStyle('C')`) on streamed sheets style every
    written cell; cells never written in that column stay default (the column
    style is also recorded for files saved without streaming).
14. **Style read-back** (updated in 4.2) — getters return local writes first,
    then the effective stylesheet state (loaded files included). Range styles
    read from the range's top-left cell; borders/protection getters remain
    local-only.
15. **Comment rich text** — comments are plain text; `Run::getFont()` throws.
16. **Auto-filter on streamed sheets** — when the auto-filter is the only
    non-streamable op, the `<autoFilter>` element is injected into the saved
    container directly (raw zip copy + one worksheet rewrite), so million-row
    filtered exports stay streaming. When other save-time ops force a degrade
    anyway, the filter rides that instead. The `_xlnm._FilterDatabase` defined
    name PhpSpreadsheet writes is omitted (Excel does not require it).
17. **Bulk calculated reads** — `toArray(calculateFormulas: true)` evaluates
    only formula cells **without a cached result** (anything Excel or
    excelize previously saved is trusted, like PhpSpreadsheet with
    pre-calculated formulas enabled). A formula whose evaluation errors comes
    back empty rather than throwing.
18. **Conditional formatting model** — color scales and data bars use the
    easy-excel `setColorScale`/`setDataBar` helpers rather than
    PhpSpreadsheet's `ConditionalColorScale`/`ConditionalDataBar` object
    graphs; range styles apply one rule list per `setConditionalStyles` call
    (replacing semantics within a range).
19. **Formula coverage** — 466/529 functions; the differences are listed in
    FORMULAS.md and unknown functions error at calculation, not at write.
20. **Gradient fill angles** — PhpSpreadsheet stores an exact rotation;
    excelize supports discrete shading directions, so the angle buckets to
    the nearest of horizontal/vertical/diagonal-up/diagonal-down (path
    gradients → from-center).
21. **Write-after-save reopens the workbook** (fixed in 4.3) — excelize
    silently discards model edits made after a StreamWriter flush, so the
    first mutation following a save of a streamed workbook triggers a
    serialize-and-reopen (correct, O(file size)). Without it the edit would
    be lost; save-then-edit-then-save flows now round-trip.
22. **Rich text applies at save** — a rich-text cell value buffers its plain
    text (so dimensions and `getValue()` are correct mid-write) and the
    formatted runs are applied to the model at save. Setting a rich-text
    value then overwriting the same cell with a plain value across the
    stream boundary is not ordering-guaranteed.
23. **Auto-filter doesn't hide rows** — like excelize (and the OOXML format),
    setting a column rule records the criteria but does not hide
    non-matching rows; Excel re-applies the filter on open. PhpSpreadsheet
    behaves the same. Column rules also accept at most two clauses joined by
    AND/OR (the OOXML custom-filter limit).
24. **Formula cache: opt-in and numeric-only** —
    `Writer\Xlsx::setPreCalculateFormulas(true)` evaluates every formula at
    save and stores the result beside it, so readers that trust the cached
    `<v>` show values rather than blanks.
    **Opt-in, deliberately.** The getter still reports PhpSpreadsheet's `true`
    default for API parity, but the pass only runs when
    `setPreCalculateFormulas()` was called explicitly: acting on the inherited
    default would read every formula back and force a streamed workbook into
    the full in-memory model — an OOM risk on million-row exports, and a
    silent trade of the property this engine exists for.
    **Numeric results only, and the check is stricter than it looks.** A text
    or boolean result cannot be cached correctly — excelize writes a
    shared-string *index* into `<v>`, which reads back as `0`/`1`. Those, and
    error results, recompute on open. `ParseFloat` alone is not sufficient to
    decide: `TEXT(A1,"0000")` yields the *string* `"0042"`, which parses as 42
    and would be cached — and rendered — as a number, losing the padding the
    formula exists to produce. excelize's calc engine knows the real type
    internally (`formulaArg.Type`) but `CalcCellValue` returns only a string,
    so a float→string→float round trip is used instead: a genuine number
    survives it unchanged, a formatted string does not.
    **Cell type is corrected at save.** excelize's `SetCellFormula` ends with an
    unconditional `c.T = "str"` and exposes no way to reset it (every value
    setter calls `removeFormula`, so writing the value afterwards deletes the
    formula). In OOXML `t="str"` means *formula returning a string*, so a
    cached number would lose its format — a `#,##0` total rendering `2500`
    instead of `2,500`, and excelize's own reader takes a
    `case "str": return c.V` path that skips formatting entirely. The saved
    container is therefore patched, exactly as streaming auto-filters already
    are. The patch is driven by the **exact cell references this save cached**,
    never by scanning for numeric-looking values, so formulas already present
    in a loaded workbook are left exactly as they were. Worksheet parts with
    nothing to change are copied without recompressing, and the passes stage
    through a temp file rather than memory. Both compose when a save needs
    both.
    Independently, `calcPr/@fullCalcOnLoad` is set by default (free, one
    attribute): spreadsheet **applications** recalculate on open regardless.
    Disable via `Native::setFullCalcOnLoad($handle, false)`.
25. **Image anchoring** — drawings use one-cell anchoring (fixed size, the
    image keeps its dimensions when rows/columns resize), matching
    PhpSpreadsheet. excelize's default two-cell anchoring (image stretches
    with the cells) is not used.
26. **`Shared\Font` takes `?object`, not `Style\Font`** — Compat's
    `Style\Font` is bound to its owning `Style` and cannot be constructed
    standalone, while PhpSpreadsheet callers pass whatever font object they
    hold. The `Shared\Font`/`Shared\Drawing` helpers therefore accept any
    object exposing `getName()`/`getSize()` and fall back to Calibri 11 for
    anything else, rather than fataling on a type mismatch.
27. **`Shared\Font` metrics are table-driven** — PhpSpreadsheet measures
    rendered text with GD/afm font metrics; easy-excel uses a lookup table for
    the common families plus a linear approximation elsewhere, consistent with
    the save-time auto-size approximation (divergence 10).

28. **Page breaks apply at save** — like other non-streamable ops, a break is
    queued and written when the workbook is flushed. `BREAK_NONE` removes a
    break at that reference rather than adding one, matching
    PhpSpreadsheet's default argument.
29. **Selection merges into panes** — excelize models selection as part of the
    pane record. easy-excel reads the sheet's current panes back and folds the
    selection in, so freeze and selection compose. Setting a selection on a
    sheet with no panes writes a pane-less selection, as Excel does.

30. **Chart axis coverage is what excelize models** — tick-mark style,
    crossing point, axis orientation beyond min/max reversal, time units and
    display units are accepted and ignored: excelize has no field for them, so
    throwing would break charts that are otherwise correct. Manual plot-area
    geometry (`Layout` x/y/w/h) and gridline line formatting are stored and
    round-trip through the getters but do not affect the rendered chart.
31. **Chart size comes from anchors, approximately** — PhpSpreadsheet anchors a
    chart between two cells; excelize sizes it in pixels. The span is converted
    with the OOXML default grid (64px per column, 20px per row), so the chart
    lands in the right place at close to the right size, not byte-identically.


## Aliasing modes

The `PhpOffice\PhpSpreadsheet\*` → `EasyExcel\Compat\*` bridge runs in one of
three modes, chosen by `aliasMode()` (`php/src/aliasing.php`) and overridable
with the `EASY_EXCEL_ALIAS` environment variable:

| Mode | When | Behaviour |
|---|---|---|
| `strict` | **default when the native extension is loaded** (or `EASY_EXCEL_ALIAS=strict`/`force`) | All-or-nothing. Implemented classes resolve to Compat; any `PhpOffice\PhpSpreadsheet\*` class Compat does **not** implement throws `EasyExcel\UnsupportedApiException`. A request is served entirely by Compat or it fails — a handle-based workbook can never be mixed with a real object graph. |
| `off` | **default when the extension is absent** (or `EASY_EXCEL_ALIAS=off`) | No aliasing; everything resolves to a real `phpoffice/phpspreadsheet` install (add it as a dependency). Use this to run on stock PhpSpreadsheet, e.g. for A/B output comparison. |
| `fallback` | `EASY_EXCEL_ALIAS=fallback` (extension required) | Hybrid escape hatch: alias what Compat implements, defer everything else to the real package per class. Convenient for incremental adoption, but can mix object models within one request — opt in knowingly. |

Strict mode throws even via a defensive `class_exists('PhpOffice\…')` probe;
that is intentional — under all-or-nothing an uncovered class is a coverage
gap to close (or a cue to switch the whole request to `off`/`fallback`), not
something to paper over silently.

**Eager binding.** In `strict`/`fallback` mode the bootstrap binds the whole
implemented surface up front (`eagerAliasCompat()`) instead of waiting for
autoload; the prepended autoloader remains only as the strict-mode tripwire
for unimplemented classes. Lazy-only aliasing had two holes: PHP never
autoloads for parameter/return/`instanceof` checks, so a Compat object hitting
consumer code type-hinted with the PhpOffice name (e.g. `function
f(Worksheet $ws)`) fataled with a TypeError unless something had referenced
the class first; and composer prepends its own autoloader, so a bootstrap
loaded before `vendor/autoload.php` lost the `PhpOffice\*` namespace to a
co-installed real package. Binding ~60 classes eagerly costs ~12 ms cold
(one-time per process). Set `EASY_EXCEL_EAGER=0` to restore lazy-only
aliasing. Names already defined (a real PhpSpreadsheet class loaded before
bootstrap) are left untouched.

**Surface diff (CI gate).** `php/tools/compat-surface-diff.php` reflects a real
PhpSpreadsheet install and reports every class/method/constant Compat is
missing. Run it against a frozen baseline so a *new* gap (e.g. a PhpSpreadsheet
version bump adding constants) fails CI instead of surfacing at runtime:

```
composer require --dev phpoffice/phpspreadsheet
php tools/compat-surface-diff.php --members                              # full report
php tools/compat-surface-diff.php --baseline=.compat-surface.json        # gate (exit 1 on new gaps)
php tools/compat-surface-diff.php --update-baseline=.compat-surface.json # bump deliberately
```

## Not supported — by design (wave 5.5)

Verified against the shipped Compat tree (`php/src/EasyExcel/Compat`): the
alias surface is derived by scanning that directory (`compatSurfaceClasses()`),
so **any name without a file there throws `UnsupportedApiException` in
`strict` mode**.

Everything below is a deliberate exclusion, not a backlog item. Each entry
says why the gap is structural and what to do instead. Waves 5.1–5.4 closed
the rest: **49 of the 53 `PhpOffice\*` names imported across the two audited
production apps now resolve under Compat**, and the four that do not are
listed here.

### The escape hatch

None of these force an all-or-nothing choice. `EASY_EXCEL_ALIAS=fallback`
aliases every class Compat implements and defers the rest to a real
`phpoffice/phpspreadsheet` install, per class — so an app can generate its
bulk exports through the native engine and keep upstream for one report that
needs raw OOXML. `EASY_EXCEL_ALIAS=off` returns the whole request to upstream.
The trade-off is documented under "Aliasing modes": `fallback` can mix object
models within a single request, so keep the boundary at the export level.

### 1. Custom OOXML writer parts

`Writer\Xlsx\WriterPart`, `Writer\Xlsx\Worksheet`, `Shared\XMLWriter`

PhpSpreadsheet builds xlsx by composing PHP writer-part classes, each
serialising a fragment of the package; subclassing one lets an app inject
arbitrary XML. excelize owns serialisation end to end — there is no part
registry to hook, and no point in the pipeline where a PHP-authored fragment
could be spliced in without re-implementing the writer in PHP, which is the
thing the native engine exists to avoid. The two models are mutually
exclusive.

`Shared\XMLWriter` exists only to serve this pattern; nothing else in the
audited apps uses it.

**Instead:** keep the affected export on upstream via `fallback`. Anything
expressible through the supported API (styles, charts, validations,
conditional formats) needs no custom part.

### 2. Subclassing `Spreadsheet` / `Worksheet`

Compat's workbook and worksheet objects are thin facades over a handle into
Go-side state — there is no PHP object graph holding cells, so a subclass has
nothing to extend or intercept. Overriding a method changes what PHP asks the
extension to do; it cannot change what the extension writes.

**Instead:** `Spreadsheet::copySheet()` covers sheet duplication (the common
reason to subclass), and the native chart/image APIs cover injection that
would otherwise be done by overriding a writer.

### 3. Chart image rendering

`Chart\Renderer\*` (incl. `JpGraph`)

easy-excel emits charts as real Excel chart parts, which Excel and LibreOffice
render themselves. Rasterising a chart to PNG in PHP is a different job,
needing a plotting library and a font stack the engine deliberately does not
carry.

**Instead:** `Settings::setChartRenderer()` is accepted and ignored (wave 5.2)
and `Chart::render()` returns `false` (wave 5.4) — the value PhpSpreadsheet
itself returns when no renderer is configured, so callers take their existing
"no image available" branch rather than fataling. Charts in the generated
xlsx are unaffected and fully rendered.

### 4. `Style\ConditionalFormatting\MergedCellStyle`

Resolves the effective style of one cell by folding in matching conditional
rules and table styles. It is not a leaf class: it needs `StyleMerger`,
`CellStyleAssessor`, `CellMatcher`, and — through
`Worksheet::getTablesWithStylesForCell()` — the whole `Worksheet\Table`
subsystem with its dxf style model, none of which Compat implements.

Earlier revisions of this file dismissed it as "reachable only from the forked
HTML writers". That was wrong: erp-add-ons still constructs one directly after
the wave-5.1 re-parenting. It stays out because it is a subsystem, not because
it is unreachable.

**Instead:** `getStyle()` returns the cell's base style, and
`getConditionalStyles()` returns the rules on a range — enough to resolve
matches manually where an app needs the merged result.

### Formats

Readers/Writers: Ods, Xls, Slk, Gnumeric — not planned for the native
engine. In `strict` mode these throw `UnsupportedApiException`; use
`fallback`/`off` with a real `phpoffice/phpspreadsheet` install, or convert
externally.

`Writer\Pdf` is supported, and keeps upstream's shape: an **abstract** base
carrying the HTML renderer and the page setup (`get/setPaperSize`,
`get/setOrientation`, `get/setFont`, `get/setTempDir`, plus
`resolvePaperSize()` / `resolveOrientation()` which fold the writer override
together with the sheet's own page setup), with the HTML→PDF step supplied by
a driver subclass. The consumer picks a driver by class, exactly as upstream:

| Driver | Requires |
| --- | --- |
| `Writer\Pdf\Mpdf` | `mpdf/mpdf` |
| `Writer\Pdf\Tcpdf` | `tecnickcom/tcpdf` |
| `Writer\Pdf\Dompdf` | `dompdf/dompdf` |
| `Writer\Pdf\Snappy` | `knplabs/knp-snappy` + a `wkhtmltopdf` binary |

`Snappy` has no counterpart upstream: the others embed a PHP rendering engine,
it shells out to wkhtmltopdf. Because the binary renders with a real browser
engine it handles CSS the PHP engines do not, so it suits heavily styled
reports; inject the Snappy instance with `setSnappy()` before saving.

The polyfill requires none of these libraries — a driver whose library is
absent throws and names the package to install. `IOFactory::createWriter()`
does not map `'Pdf'`, again as upstream, since the base is abstract.
