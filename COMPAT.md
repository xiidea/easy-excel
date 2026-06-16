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
| IOFactory | `createWriter/Reader` (Xlsx, Csv), `load`, `identify` | |
| Writer\Xlsx | `save` (paths and `php://` streams) | |
| Writer\Csv | `set/getDelimiter`, `setEnclosure` (only `"`), `set/getLineEnding`, `set/getUseBOM`, `set/getSheetIndex`, `save` | plus `setSanitizeFormulas()` (easy-excel extra, opt-in OWASP guard) |
| Reader\Xlsx | `load`, `setReadDataOnly`, `canRead` | |
| Reader\Csv | `load`, `setDelimiter`, `setEnclosure`, `setSheetIndex`, `canRead` | streams in 1k-row chunks |
| Value binding | DefaultValueBinder semantics: numeric strings → numbers (leading-zero strings preserved), `=…` → formula, `DateTimeInterface` → Excel serial | |

## Supported (Phase 2 — formatting & structure)

| Area | API | Notes |
|---|---|---|
| Style | `getStyle(range)->applyFromArray`, `getFont` (bold/italic/size/name/underline/strikethrough/color/super/subscript), `getFill` (pattern fills + start/end color), `getBorders` (top/bottom/left/right/allBorders/outline, style + color), `getAlignment` (horizontal/vertical/wrapText/shrinkToFit/textRotation/indent), `getProtection` (locked/hidden), `getNumberFormat` | partial styles layer in application order like PhpSpreadsheet's supervisor; styles applied **before** their rows are written ride the StreamWriter at zero cost |
| Style helpers | `Color` (ARGB/RGB + constants), all `Border::BORDER_*`, `Fill::FILL_*`, `Alignment::HORIZONTAL_*/VERTICAL_*`, `Protection::PROTECTION_*`, `NumberFormat::FORMAT_*` constants | |
| Dimensions | `getColumnDimension(+ByColumn)->setWidth/setAutoSize`, `getRowDimension->setRowHeight` | auto-size is approximated at save time (divergence 10) |
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

## Documented divergences

1. **`toArray(formatData: false)` types** — values come back from excelize as
   strings and are cast with `is_numeric()`. Text cells that *look* numeric
   (e.g. `"1e3"` stored explicitly as a string) come back as numbers, where
   PhpSpreadsheet preserves them. Explicitly-typed strings written in the
   same session are safe; re-loaded files lose that distinction.
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
24. **No pre-computed formula cache** — formula cells are written with the
    formula but without a cached `<v>` result (PhpSpreadsheet pre-calculates
    and stores it). Excel, LibreOffice and `getCalculatedValue()` recompute
    on open and display the correct value; headless readers that trust the
    cache without recalculating see a blank until they evaluate. excelize has
    no recalculate-and-store-all step, so this is inherent to the engine.
25. **Image anchoring** — drawings use one-cell anchoring (fixed size, the
    image keeps its dimensions when rows/columns resize), matching
    PhpSpreadsheet. excelize's default two-cell anchoring (image stretches
    with the cells) is not used.

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

## Not yet supported (throws a clear exception)

- Gradient fills, diagonal/vertical/horizontal borders
- PhpSpreadsheet's `Chart` object model (`PhpOffice\PhpSpreadsheet\Chart\*`):
  use the native declarative API (`Worksheet::addNativeChart`) instead
- Workbook encryption / password-protected open
- Readers/Writers: Ods, Xls, Html, Pdf, Slk, Gnumeric — not planned for the
  native engine. In `strict` mode (the default with the extension) these throw
  `UnsupportedApiException`; set `EASY_EXCEL_ALIAS=off` (or `fallback`) and
  install the real `phpoffice/phpspreadsheet` to handle them, or convert
  externally
- Custom value binders (`Cell::setValueBinder`), read filters with PHP
  callbacks — planned via declarative equivalents (PHP callbacks across the
  CGO boundary are the documented slow path)
- `Worksheet::getRowIterator()/getColumnIterator()` (`toArray` chunked reads
  cover most uses)
