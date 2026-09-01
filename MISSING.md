# Not implemented

PhpSpreadsheet APIs the polyfill does **not** provide. COMPAT.md documents
what *is* supported and where supported behaviour intentionally diverges.
Calling anything below fails loudly (class not found / clear exception) —
never with a silently different file.

**Phase 5 is complete.** Waves 5.1–5.5 (PLAN.md §13) closed the gaps found by
auditing two production Symfony report apps against the shim: **49 of the 53
`PhpOffice\*` names those apps import now resolve under Compat**. What remains
is deliberate — see "By design" below and the matching section in COMPAT.md,
which carries the full rationale for each.

## By design (will not be implemented)

These are structural mismatches with a Go/excelize engine, not backlog items.
`EASY_EXCEL_ALIAS=fallback` defers any of them to a real
`phpoffice/phpspreadsheet` install per class, so an app can keep its bulk
exports native and route one report to upstream.

- **Custom OOXML writer parts** — `Writer\Xlsx\WriterPart`,
  `Writer\Xlsx\Worksheet`, `Shared\XMLWriter`. excelize owns serialisation
  end to end; there is no part registry to hook (COMPAT.md §1)
- **Subclassing `Spreadsheet` / `Worksheet`** — Compat objects are handle
  facades over Go state, with no PHP object graph to extend (COMPAT.md §2)
- **Chart image rendering** — `Chart\Renderer\*` (incl. `JpGraph`). Charts
  are emitted as native Excel chart parts; `Chart::render()` returns false so
  callers take their no-image branch (COMPAT.md §3)
- **`Style\ConditionalFormatting\MergedCellStyle`** — needs `StyleMerger`,
  `CellStyleAssessor`, `CellMatcher` and the whole `Worksheet\Table` subsystem
  (COMPAT.md §4)
- **`Writer\Html` subclassing** — the Compat writer is an independent pure-PHP
  renderer, not a port, so overriding its protected internals does not
  compose. The public writer API is supported; wave 5.1 showed the audited
  apps never needed the inheritance at all
- Readers/writers for Ods, Xls, Slk, Gnumeric
- **`getCellCollection()`** — cell data lives in Go, not in a PHP collection

## Open gaps (no observed consumer yet)

Not exercised by either audited app, so unprioritised rather than refused:

- Auto-filter **column rule** introspection (range getter landed in 4.2)
- `removeConditionalStyles`
- `clone $sheet` / `Spreadsheet::addExternalSheet` (use
  `Spreadsheet::copySheet` instead)
- Vertical/horizontal borders (conditional-formatting-only border sides)
- Header/footer images, cell background images (file & memory drawings
  anchored to cells are supported)
- 63 of PhpSpreadsheet's 529 calculation functions (list in FORMULAS.md)

## Behavioural divergences worth planning around

Not missing APIs — these produce wrong-looking output even once every class
exists, so they outrank the open gaps above for anyone migrating a
report-heavy app:

- **Formula cache is numeric-only** — *largely fixed*. Pre-calculation is now
  wired to `Writer\Xlsx::setPreCalculateFormulas()` and defaults on, matching
  PhpSpreadsheet, and `calcPr/@fullCalcOnLoad` is set by default. What remains:
  excelize cannot store a **text or boolean** formula result correctly (it
  writes a shared-string index into `<v>`), so those still recompute on open.
  Numeric results — the overwhelming majority in the audited reports — are
  cached (COMPAT.md §24)
- **Style-after-write degrade** — styling rows *after* writing them queues the
  work and triggers the one-time serialize-and-reopen at save, forfeiting the
  streaming win. Both audited apps do this for subtotal/total rows. No API is
  missing; the throughput claim just does not hold for that pattern
  (COMPAT.md §9)
- **Auto-filter does not hide rows** — column rules are recorded; Excel
  re-applies on open (COMPAT.md §23)
- **`Calculation` array-formula toggles** — the cache controls are accepted
  no-ops since wave 4.1; calculation is delegated to excelize

## Wave history

Phase 5 (2026-08-19), scoped from the two-app audit:

| Wave | Closed |
|---|---|
| 5.1 | `Shared\StringHelper`; consumer-side re-parenting of both `HTMLWriter`s from `Writer\Html` to `BaseWriter` |
| 5.2 | `Writer`/`Reader`/`Calculation` exceptions, `Reader\IReader`, `Settings`, `Cell\CellAddress`+`AddressRange`, `Worksheet\BaseDrawing`, `Shared\File`+`Font`+`Drawing` |
| 5.3 | `setBreak()`, `setSelectedCells()`, `calculateColumnWidths()` |
| 5.4 | Chart axis model: `Chart\Axis`, `GridLines`, `Layout`, `ChartColor`, constructor axis params, `DataSeries::EMPTY_AS_*` |
| 5.5 | Documentation of the by-design exclusions above |

Phase 4 (2026-06-13) closed the probe-driven gaps in four waves: value
binders, document properties, print layout, encryption, gradient fills and
diagonal borders (4.1); iterators, read filters, style read-back and
introspection (4.2); row/column/sheet structure editing, sheet views,
headers/footers and margins (4.3); rich text, memory drawings, the
`Chart\*` object model and auto-filter column rules (4.4).

## Verified against PhpSpreadsheet

`data/public/index.php` is a self-verifying probe: it generates the ERP
report through the `PhpOffice\PhpSpreadsheet\*` aliases, reloads it, and
diffs the data table against `data/public/phpspreadsheet.xlsx` (the same
report from real PhpSpreadsheet). Run it over HTTP (see `http-verify.sh`):

```
docker run -d -e SERVER_NAME=":80" -v $PWD/data:/app \
  -v $PWD/php:/opt/easy-excel/php -p 8085:80 frankenphp-easy-excel
curl http://localhost:8085/        # → REPORT TEST PASS
```

Want one of these? Open an issue at
[xiidea/easy-excel](https://github.com/xiidea/easy-excel/issues) — gaps get
prioritized by real-world usage, and this file shrinks as they land.
