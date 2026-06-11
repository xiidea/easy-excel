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
| Style | `getStyle()->getNumberFormat()->set/getFormatCode`, `applyFromArray(['numberFormat' => …])` | number formats only in Phase 1 |
| Value binding | DefaultValueBinder semantics: numeric strings → numbers (leading-zero strings preserved), `=…` → formula, `DateTimeInterface` → Excel serial | |

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
4. **Streaming degrade** — out-of-order writes, reads, or number formats on a
   sheet with already-streamed rows trigger a one-time serialize-and-reopen
   of the workbook (correct, but O(file size)). Sequential writers never hit it.
5. **CSV enclosure** — only `"` is supported; `setEnclosure` with anything
   else throws.
6. **`createSheet($index)`** — only appending (index = count or null);
   arbitrary insert positions throw.
7. **`setReadDataOnly`** — accepted for API parity; reads are already
   values-only fast paths.
8. **Number-format rendering** — formatted values are rendered by excelize;
   rare locale-specific format codes may render differently from
   PhpSpreadsheet. Differences found by the test suite get fixed or listed here.

## Not yet supported (throws a clear exception)

- Styles beyond number formats: Font, Fill, Borders, Alignment, Protection (Phase 2)
- Auto-filter, freeze panes, column widths/row heights, hyperlinks, comments,
  defined names, page setup (Phase 2)
- Charts, data validation, conditional formatting, images/drawings,
  protection/encryption (Phase 3)
- Readers/Writers: Ods, Xls, Html, Pdf, Slk, Gnumeric — not planned for the
  native engine; install the real `phpoffice/phpspreadsheet` alongside (the
  alias bootstrap then defers to it) or convert externally
- Custom value binders (`Cell::setValueBinder`), read filters with PHP
  callbacks — Phase 2 via declarative equivalents (PHP callbacks across the
  CGO boundary are the documented slow path)
- `Worksheet::getRowIterator()/getColumnIterator()` (Phase 2; `toArray`
  chunked reads cover most uses)
