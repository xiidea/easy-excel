# Not implemented

PhpSpreadsheet APIs the polyfill does **not** provide. Found by running real
PhpSpreadsheet code against it (`data/public/index.php` is one such probe);
COMPAT.md documents what *is* supported and where supported behavior
intentionally diverges. Calling anything below fails loudly (class not
found / clear exception) — never with a silently different file.

**An implementation plan for closing these gaps exists**: PLAN.md §13
"Phase 4 — compat completion" orders everything below into four ROI waves
with verified excelize APIs and effort estimates. Items land there and get
deleted here.

## Found by the ERP report probe (`data/public/index.php`)

Closed by wave 4.3 (2026-06-13): insert/remove rows and columns,
`createSheet($index)`, sheet copy (`Spreadsheet::copySheet` extra), sheet
views (gridlines/zoom/RTL/tab color), headers/footers, page margins — plus
a correctness fix: post-save mutations were silently dropped by excelize
on stream-flushed sheets; they now reopen first (COMPAT.md §21).

Closed by wave 4.2 (2026-06-13): `getDefaultStyle()`, row/column iterators,
`IReadFilter`, style read-back from loaded files + `duplicateStyle`,
validation/conditional/defined-name/auto-filter getters.

Closed by wave 4.1 (2026-06-13): custom value binders, document properties
(`getProperties()`; `setManager` is kept PHP-side only — excelize has no
field for it), print titles + print area, the `getConditionalStyles()`
getter, workbook encryption (writer/reader `setPassword()`, easy-excel
extras), gradient fills, diagonal borders, `unmergeCells` + merge getter,
and calculation-cache no-ops.

## Known gaps (by area)

**Reading / introspection**
- `getCellCollection()` / existing-cells-only iteration flags
- Auto-filter **column rule** introspection (range getter landed in 4.2)

**Structure editing**
- `removeConditionalStyles`
- `clone $sheet` / `Spreadsheet::addExternalSheet` (use
  `Spreadsheet::copySheet` instead)

**Content types**
- `RichText` as a **cell value** (multi-format runs in one cell; comments
  accept plain text only and `Run::getFont()` throws)
- PhpSpreadsheet's `Chart\*` object model — use the native declarative API
  instead (`Worksheet::addNativeChart`, NATIVE.md)
- Vertical/horizontal borders (conditional-formatting-only border sides)
- Drawings beyond file-based images (memory drawings, headers/footers
  images, cell background images)

**Formats & security**
- Readers/writers: Ods, Xls, Html, Pdf, Slk, Gnumeric — install the real
  `phpoffice/phpspreadsheet` alongside (the alias bootstrap stays dormant
  and defers to it) or convert externally
- 63 of PhpSpreadsheet's 529 calculation functions (list in FORMULAS.md)

**Misc**
- `Calculation` array-formula toggles (the cache controls are accepted
  no-ops since wave 4.1) — calculation is delegated to excelize
- Cell autofilter object model (`getAutoFilter()->setRange()`,
  column rules) — only `setAutoFilter($range)`

Want one of these? Open an issue at
[xiidea/easy-excel](https://github.com/xiidea/easy-excel/issues) — gaps get
prioritized by real-world usage, and this file shrinks as they land.
