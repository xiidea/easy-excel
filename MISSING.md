# Not implemented

PhpSpreadsheet APIs the polyfill does **not** provide. Found by running real
PhpSpreadsheet code against it (`data/public/index.php` is one such probe);
COMPAT.md documents what *is* supported and where supported behavior
intentionally diverges. Calling anything below fails loudly (class not
found / clear exception) — never with a silently different file.

## Found by the ERP report probe (`data/public/index.php`)

| API | Status / workaround |
|---|---|
| `Cell::setValueBinder()`, `DefaultValueBinder`, `IValueBinder`, custom binder classes | Not supported — binding happens in Go with DefaultValueBinder semantics (PHP callbacks across the CGO boundary are the documented slow path, PLAN.md §5). Workaround: `setCellValueExplicit()` for values that need a non-default type, e.g. >2⁵³ IDs as strings. |
| `Spreadsheet::getDefaultStyle()` | Workbook-wide default font/style not supported. Workaround: apply the style to the used range. |
| `Spreadsheet::getProperties()` (`setTitle`, `setSubject`, `setCreator`, …) | Document metadata not written. |
| `PageSetup::setRowsToRepeatAtTopByStartAndEnd()` (print titles, print area) | Page setup covers orientation / paper size / fit-to-page only. |
| `Style::getConditionalStyles()` (getter) | Rules are write-only; build the rule array locally and call `setConditionalStyles()` once. |

## Known gaps (by area)

**Reading / introspection**
- Style read-back from loaded files (`getFont()->getBold()` etc. return what
  was set on that PHP object, not stylesheet state — COMPAT.md §14)
- Merge/auto-filter/validation/conditional getters
- `getRowIterator()` / `getColumnIterator()` / `getCellCollection()`
  (chunked `toArray` covers bulk reads)
- `Reader\IReadFilter` with PHP callbacks
- Defined-name listing (`getDefinedNames()`)

**Structure editing**
- `insertNewRowBefore` / `removeRow` / `insertNewColumnBefore` /
  `removeColumn`
- `unmergeCells`, `duplicateStyle`, `removeConditionalStyles`
- Sheet copy/clone (`Spreadsheet::addExternalSheet`, `Worksheet::copy`),
  `createSheet($index)` at arbitrary positions
- Sheet views: gridline toggle, tab color, zoom, right-to-left

**Content types**
- `RichText` as a **cell value** (multi-format runs in one cell; comments
  accept plain text only and `Run::getFont()` throws)
- PhpSpreadsheet's `Chart\*` object model — use the native declarative API
  instead (`Worksheet::addNativeChart`, NATIVE.md)
- Gradient fills; diagonal/vertical/horizontal borders
- Drawings beyond file-based images (memory drawings, headers/footers
  images, cell background images)

**Formats & security**
- Readers/writers: Ods, Xls, Html, Pdf, Slk, Gnumeric — install the real
  `phpoffice/phpspreadsheet` alongside (the alias bootstrap stays dormant
  and defers to it) or convert externally
- Workbook encryption / password-protected open (sheet protection **is**
  supported)
- 63 of PhpSpreadsheet's 529 calculation functions (list in FORMULAS.md)

**Misc**
- `Calculation` engine controls (`disableCalculationCache`, array-formula
  toggles) — calculation is delegated to excelize
- Headers/footers, page margins, print options beyond page setup
- Cell autofilter object model (`getAutoFilter()->setRange()`,
  column rules) — only `setAutoFilter($range)`

Want one of these? Open an issue at
[xiidea/easy-excel](https://github.com/xiidea/easy-excel/issues) — gaps get
prioritized by real-world usage, and this file shrinks as they land.
