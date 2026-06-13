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

Closed by wave 4.4 (2026-06-13): rich-text cell values with per-run fonts,
GD `MemoryDrawing`, the PhpSpreadsheet `Chart\*` object model
(`Worksheet::addChart`), and auto-filter column rules
(`getAutoFilter()->getColumn()`). This completes Phase 4 — MISSING.md now
lists only items that stay out by design.

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
- Vertical/horizontal borders (conditional-formatting-only border sides)
- Header/footer images, cell background images (file & memory drawings
  anchored to cells are supported)

**Formats & security**
- Readers/writers: Ods, Xls, Html, Pdf, Slk, Gnumeric — install the real
  `phpoffice/phpspreadsheet` alongside (the alias bootstrap stays dormant
  and defers to it) or convert externally
- 63 of PhpSpreadsheet's 529 calculation functions (list in FORMULAS.md)

**Misc**
- `Calculation` array-formula toggles (the cache controls are accepted
  no-ops since wave 4.1) — calculation is delegated to excelize
- Auto-filter does not hide non-matching rows (column rules are recorded;
  Excel re-applies on open — COMPAT.md §23)

Want one of these? Open an issue at
[xiidea/easy-excel](https://github.com/xiidea/easy-excel/issues) — gaps get
prioritized by real-world usage, and this file shrinks as they land.
