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
11. **Auto-filter, hyperlinks, comments, auto-size, page setup** — excelize
    cannot stream these, so they are applied at save (degrading once if the
    sheet streamed). The data path itself stays streaming.
12. **Range styles assume uniform ranges** — `getStyle('A1:C10')` applies one
    merged style to the whole range. Earlier styles fully containing the
    range are layered in (like PhpSpreadsheet); partially-overlapping earlier
    styles are not re-read per cell.
13. **Full-column styles** (`getStyle('C')`) on streamed sheets style every
    written cell; cells never written in that column stay default (the column
    style is also recorded for files saved without streaming).
14. **Style read-back** — `getFont()->getBold()` etc. return what was set on
    that PHP style object, not the stylesheet state of a loaded file.
15. **Comment rich text** — comments are plain text; `Run::getFont()` throws.

## Not yet supported (throws a clear exception)

- Gradient fills, diagonal/vertical/horizontal borders
- Charts, data validation, conditional formatting, images/drawings,
  protection/encryption (Phase 3)
- Readers/Writers: Ods, Xls, Html, Pdf, Slk, Gnumeric — not planned for the
  native engine; install the real `phpoffice/phpspreadsheet` alongside (the
  alias bootstrap then defers to it) or convert externally
- Custom value binders (`Cell::setValueBinder`), read filters with PHP
  callbacks — planned via declarative equivalents (PHP callbacks across the
  CGO boundary are the documented slow path)
- `Worksheet::getRowIterator()/getColumnIterator()` (`toArray` chunked reads
  cover most uses)
