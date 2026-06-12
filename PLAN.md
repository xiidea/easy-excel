# easy-excel — PhpSpreadsheet-compatible PHP extension in Go for FrankenPHP

**Status:** PLAN — awaiting approval before any code is written.
**Author:** Principal Architecture review, 2026-06-11

---

## 1. Problem statement

PhpSpreadsheet is the de-facto PHP spreadsheet library, but it is slow and memory-hungry:
every cell is a PHP object (or a packed entry in a cell collection), styles are PHP object
graphs, and XLSX writing builds large XML strings in PHP userland. Generating a 1M-row
export routinely needs gigabytes of RAM and minutes of CPU. Streaming alternatives
(OpenSpout, fast-excel-writer) fix memory but sacrifice the rich API and are still bound by
PHP's per-cell interpretation cost.

**Goal:** a FrankenPHP native extension, written in Go, that exposes a
PhpSpreadsheet-compatible API while doing all heavy lifting (XML generation, ZIP
compression, parsing, formula evaluation) in compiled Go — targeting **high throughput with
minimal, constant memory**.

---

## 2. Hard constraints discovered (these shape everything)

From the FrankenPHP extension documentation (https://frankenphp.dev/docs/extensions/):

| # | Constraint | Architectural consequence |
|---|-----------|---------------------------|
| C1 | Go-backed PHP classes are **opaque**: objects cannot be used as method parameters or return types; no property access from PHP | We cannot return a `Worksheet` object from `Spreadsheet::getActiveSheet()` at the extension level. → **Handle-based core + PHP userland shim** (see §4). |
| C2 | **CGO call overhead** is fixed and non-trivial (~50–200ns + marshaling per call) | Per-cell extension calls are forbidden in the hot path. → **Row/range batch API**, PHP shim buffers writes. |
| C3 | String/array conversions allocate and traverse on every boundary crossing | Cross the boundary with **packed arrays of rows**, not maps; avoid round-tripping data that Go already holds. |
| C4 | **No ZTS**; FrankenPHP worker mode keeps PHP processes alive; goroutines may outlive a PHP call but not safely a request | Go side owns all shared state behind its own synchronization; per-request handle cleanup hooks mandatory (leak prevention in long-lived workers). |
| C5 | Extension generator supports scalars, arrays, nullable scalars; callables via `frankenphp.CallPHPCallable()` (expensive) | PHP callbacks (value binders, read filters) supported but documented as slow path; provide declarative Go-side equivalents. |

These constraints mean a *naive* 1:1 port of PhpSpreadsheet's API into `//export_php`
methods would be **slower** than PhpSpreadsheet for cell-by-cell code. The architecture
below is designed so idiomatic PhpSpreadsheet code still gets fast, and bulk APIs
(`fromArray`, `toArray`, writers/readers) get *dramatically* fast.

---

## 3. Library selection (build vs. buy)

| Option | Verdict | Rationale |
|--------|---------|-----------|
| **`github.com/xuri/excelize/v2`** (qax-os/excelize) | ✅ **Primary engine** | Mature (18k★), actively maintained, Apache-2.0. Has **StreamWriter** (constant-memory row streaming, ~19% faster SetRow since v2.7) and streaming row iterator for reads. Supports styles, merged cells, charts, pivot tables, data validation, defined names, rich strings, and a **calculation engine with ~535 functions** — the closest Go analogue to PhpSpreadsheet's Calculation engine. |
| `tealeg/xlsx` | ❌ | Loads whole workbook in memory; no streaming; slower development pace. |
| `plandem/xlsx` | ❌ | Unmaintained since 2021. |
| Hand-rolled SAX XML writer | ⚠️ **Phase-3 escape hatch only** | excelize's StreamWriter is already near I/O-bound; a bespoke writer (direct `xml` byte emission + `klauspost/compress` zip) buys maybe 1.5–2× more on writes but costs correctness risk. Only if benchmarks show StreamWriter is the bottleneck. |
| `klauspost/compress` | ✅ Supporting | Drop-in faster DEFLATE for the output ZIP (~2× faster than stdlib at same ratio). excelize allows custom zip via post-processing; if not cleanly pluggable, we re-compress the container at save time only if profiling justifies it. |
| `golang.org/x/sync/semaphore` | ✅ Supporting | Weighted semaphore for global concurrency control (§7). |

**Decision: excelize/v2 as the engine; do not fork it; wrap it.** Incompatibilities with
PhpSpreadsheet semantics are absorbed in our Go adapter layer, never pushed to PHP.

---

## 4. Architecture

```
┌────────────────────────────────────────────────────────────────────┐
│ PHP userland (Composer package: easy-excel/polyfill)               │
│                                                                    │
│  PhpOffice\PhpSpreadsheet\* — drop-in shim classes                 │
│  Spreadsheet, Worksheet, Cell, Style, IOFactory, Writer\Xlsx, …    │
│  • hold int64 $handle + cursor state only (tiny objects)           │
│  • buffer per-cell writes into row batches (write-behind cache)    │
│  • flush batches to extension at N rows / on read / on save        │
└───────────────┬────────────────────────────────────────────────────┘
                │  flat functions: easy_excel_*(handle, packed arrays)
┌───────────────▼────────────────────────────────────────────────────┐
│ FrankenPHP Go extension (this repo)                                │
│                                                                    │
│  bridge/    //export_php:function layer — thin, no logic           │
│             marshals packed arrays ↔ Go slices, returns handles    │
│  registry/  handle table: map[int64]*Workbook, sharded RWMutex,    │
│             request-scoped ownership, shutdown sweep (C4)          │
│  core/      Workbook/Sheet adapters over excelize                  │
│             • write path: StreamWriter when access pattern allows  │
│             • random-access path: excelize in-memory model         │
│             • style interner (hash → excelize styleID) (§6)        │
│  compat/    PhpSpreadsheet semantics: A1/R1C1 addressing, value    │
│             binders, number-format translation, date epoch (1900/  │
│             1904), formula dialect mapping                         │
│  limits/    weighted semaphore, memory budget, timeouts (§7)       │
│  io/        file/stream sinks, temp-file spill, path policy (§8)   │
└────────────────────────────────────────────────────────────────────┘
```

### 4.1 Why a PHP shim instead of `//export_php:class` everywhere

Constraint C1: extension classes are opaque and can't reference each other. PhpSpreadsheet
code is object-graph-heavy (`$sheet->getCell('A1')->getStyle()->getFont()->setBold(true)`).
The shim gives us:

- **Full API fidelity** — same namespaces, same fluent chains, `instanceof` works, existing
  code runs with only a Composer `replace` of `phpoffice/phpspreadsheet`.
- **The batching layer** (C2): `setCellValue()` appends to a PHP array; one CGO call flushes
  512 rows. Without this, per-cell CGO calls would be the #1 bottleneck.
- **Graceful degradation**: if the extension is absent, the shim can fall back to real
  PhpSpreadsheet (CI-friendly, easy adoption).

The extension itself exposes a **flat, stable C-like ABI** (~60 functions, handle-first),
which is also independently usable for maximum-performance code that skips the shim.

### 4.2 Dual write paths (the key throughput decision)

PhpSpreadsheet allows random access (`setCellValue('A1', …)` then `('A1000000', …)` then
back to `'B5'`). excelize's StreamWriter requires strictly ascending rows. We auto-select:

- **Streaming mode** (default for fresh workbooks): rows flushed in ascending order go to
  `StreamWriter` → constant ~MBs of memory regardless of row count. Out-of-order write
  triggers a one-time documented fallback.
- **Random-access mode**: excelize's normal model (still far lighter than PHP objects).
  Used automatically for loaded files and out-of-order writes.

Mode selection is transparent; a `EasyExcel\Hints::sequentialWrites()` opt-in pins
streaming mode and makes out-of-order writes throw instead of silently de-optimizing.

---

## 5. PhpSpreadsheet API compatibility matrix (phased)

Full surface is enormous (~900 classes). Phasing by real-world usage frequency:

**Phase 1 — the 90% use case (MVP):**
`Spreadsheet`, `Worksheet` (create/rename/delete/iterate), `setCellValue[Explicit]`,
`getCell`/`getValue`/`getCalculatedValue`/`getFormattedValue`, `fromArray`/`toArray`/
`rangeToArray`, `IOFactory`, `Writer\Xlsx`, `Writer\Csv`, `Reader\Xlsx` (+ `setReadDataOnly`,
`ReadFilter`), `Reader\Csv`, `Coordinate` helpers, `Cell\DataType`, default value binder.

**Phase 2 — formatting & structure:**
`Style` graph (Font, Fill, Borders, Alignment, NumberFormat, Protection), `applyFromArray`,
`getStyle(range)`, column widths/row heights, auto-size, merged cells, auto-filter, freeze
panes, hyperlinks, comments, defined names, page setup, `Date` shared helpers
(Excel epoch 1900/1904, `PhpSpreadsheet\Shared\Date::excelToDateTimeObject`).

**Phase 3 — advanced:**
`getCalculatedValue` full formula engine (delegate to excelize's 535-function engine;
publish a documented function-coverage table vs PhpSpreadsheet's ~400), charts (excelize
chart API mapped from PhpSpreadsheet chart model), data validation, conditional formatting,
images/drawings, protection/encryption (excelize supports password open/protect),
`Reader\Ods`/`Writer\Ods` (excelize: no — mark **unsupported**, fall back to real
PhpSpreadsheet via shim), `Xls` legacy (same fallback).

**Compatibility policy (per your requirement):** never introduce incompatible behavior
silently. Where Go semantics differ (formula edge cases, number-format rendering), we (a)
match PhpSpreadsheet in the `compat/` layer when cheap, (b) document divergence in
`COMPAT.md` when not, (c) keep the PhpSpreadsheet-fallback escape hatch per class. The
shim's test suite **runs relevant PhpSpreadsheet unit tests against our implementation** —
compatibility measured, not asserted.

---

## 6. Caching strategy

| Cache | Where | Why | Policy |
|-------|-------|-----|--------|
| **Style interner** | Go, per-workbook | PhpSpreadsheet code calls `applyFromArray` per cell; styles repeat massively. Hash the canonical style struct → reuse excelize styleID. Mirrors PhpSpreadsheet's own style dedup but at Go speed. | map[uint64]int, lives with workbook handle |
| **Row write-behind buffer** | PHP shim | Amortize CGO crossings (C2): flush every 512 rows or 1MB, configurable | flushed on read of dirty range, save, destruct |
| **Read row cache** | PHP shim | `toArray`/iterators fetch 1k-row chunks per CGO call, serve cells from the chunk | invalidated on write to range |
| **Number-format render cache** | Go, process-wide | Format-code → compiled formatter (excelize does some of this; we memoize the compat translation layer) | bounded LRU, 4k entries |
| **Shared-string policy** | Go | For write path default to **inline strings** in streamed sheets (excelize StreamWriter default; faster, no global table contention, slightly larger file); honor shared strings on read | per-workbook toggle |
| **Template cache** (opt-in) | Go, process-wide | Worker mode (C4): repeatedly loading the same .xlsx template re-parses XML. Cache parsed template by (path, mtime, size), clone per request | bounded LRU by memory budget, explicit `EasyExcel\Templates::warm()` |

Explicitly **not** caching: rendered output files (that's the app's/CDN's job), formula
results across requests (correctness risk).

---

## 7. Concurrency, queues, and load control

FrankenPHP worker mode = many concurrent PHP requests in one process sharing our Go
runtime. Uncontrolled, 200 concurrent 500k-row exports would exhaust memory. Design:

1. **Weighted admission semaphore** (`x/sync/semaphore`) around expensive operations
   (load, save, toArray of large ranges). Weight = estimated cost (declared size hints or
   file size). Default capacity: `max(2, NumCPU)` heavy ops; configurable via
   `EASY_EXCEL_MAX_CONCURRENT` / php.ini-style options.
2. **Bounded wait with deadline**: callers block up to `acquire_timeout` (default 30s) then
   receive a typed PHP exception (`EasyExcel\Overloaded`) the app can convert to HTTP 429 /
   queue dispatch. **Backpressure is surfaced, not hidden.**
3. **No internal job queue in v1.** PHP ecosystems already have queues (Symfony Messenger,
   Laravel queues). An internal async queue duplicates that and complicates delivery
   semantics. Instead: the blocking API + `Overloaded` signal composes with existing queues.
   *(Phase 3 option: `saveAsync()` returning a ticket + `await`, using goroutines — the
   FrankenPHP goroutine facility makes this cheap to add later.)*
4. **Per-workbook confinement**: a workbook handle is owned by one request; Go side guards
   with a per-handle mutex anyway (cheap, prevents UB if PHP code leaks handles across
   fibers). Registry itself: sharded `sync.RWMutex` (16 shards) — no global contention.
5. **Memory budget**: process-wide `atomic` accounting of live workbook estimated bytes;
   exceeding `EASY_EXCEL_MEMORY_BUDGET` (default 512MB) makes new heavy admissions fail
   fast with `Overloaded`. Combined with streaming mode, steady-state per-export memory is
   ~2–8MB, so the budget is a circuit breaker, not a daily limiter.
6. **Request-shutdown sweep** (C4): RSHUTDOWN hook releases all handles registered by that
   request — leaks impossible even if PHP code forgets `disconnect()`. Finalizers as
   backstop, not primary mechanism.

---

## 8. Security

- **Path policy**: extension-level allowlist of base directories for load/save
  (`EASY_EXCEL_ALLOWED_PATHS`, default: respect PHP `open_basedir`); reject `..`
  traversal after `filepath.Clean`; no URL schemes in v1 (no SSRF surface).
- **Zip-bomb / decompression limits on read**: cap uncompressed size, entry count, XML
  nesting depth, and total cell count (configurable; excelize already has `Options{
  UnzipSizeLimit, UnzipXMLSizeLimit }` — we set safe defaults of 1GB/16MB instead of
  unlimited).
- **XXE**: Go's `encoding/xml` does not resolve external entities — structurally immune;
  add a regression test anyway (malicious DTD fixture).
- **Formula injection (CSV)**: `Writer\Csv` gets PhpSpreadsheet-parity escaping plus an
  opt-in OWASP guard (prefix `'` for `=+-@` leading cells) — same flag name semantics.
- **Handles are capabilities**: random 64-bit handles (not sequential) + request-scoped
  ownership check → a forged/stale handle from another request errors, never aliases.
- **No PHP code execution from documents**: we never eval; PHP callbacks only run when the
  app explicitly registers them (C5).
- **Resource limits double as DoS protection** (§7: semaphore, memory budget, timeouts).
- **Supply chain**: go.mod pinned, `govulncheck` in CI, vendored builds for release
  artifacts.

---

## 9. Performance bottleneck analysis (ranked) and ROI plan

### Bottlenecks identified

| # | Bottleneck | Where | Severity |
|---|-----------|-------|----------|
| B1 | **Per-cell CGO crossings** | PHP↔Go boundary | Fatal if unmitigated — would make us *slower* than PhpSpreadsheet for chatty code |
| B2 | **PHP-side cell-object churn** | PhpSpreadsheet itself | The reason this project exists; eliminated by design (cells live in Go) |
| B3 | **String marshaling** (`GoString`/`PHPString` alloc per value) | boundary | High at 10⁷ cells; mitigated by packed-array batches (one traversal) and avoiding round-trips |
| B4 | **ZIP DEFLATE CPU** | save path | ~30–50% of save wall-time; `klauspost/compress`, compression level tunable (level 1 for API responses, 6 for archives) |
| B5 | **Non-streamed writes buffering whole sheet** | random-access mode | Solved by dual-path (§4.2): streaming default |
| B6 | **Blocking file I/O on save/load** | io/ | Acceptable (PHP request model is blocking); mitigate with buffered writers (1MB), `O_TMPFILE`+rename for atomicity; async ticket API deferred to Phase 3 |
| B7 | **Template re-parse per request** | worker mode | Template LRU cache (§6) |
| B8 | **Style lookup per cell** | core/ | Style interner (§6); O(1) hash vs PhpSpreadsheet's md5-of-serialized-object |
| B9 | **GC pressure from row slices** | Go core | `sync.Pool` for row buffers; reuse `[]interface{}` across flushes |
| B10 | **Network/large payloads** | n/a in v1 | No network I/O in scope; outputs go to disk/stream. Browser-bound payloads: recommend `Content-Encoding` passthrough (xlsx is already deflated — never double-compress; document it) |

### ROI ordering (what gets built/optimized first)

**Tier 1 — high impact / low risk (do first):**
1. Handle architecture + batched flat ABI (kills B1, B3 by design — not an optimization, a precondition).
2. excelize StreamWriter write path (kills B2, B5).
3. PHP shim write-behind row buffering (B1).
4. Request-shutdown handle sweep (correctness under load; enables worker mode at all).

**Tier 2 — high impact / medium risk:**
5. Style interner (B8) — medium risk only because style canonicalization must match PhpSpreadsheet semantics.
6. `klauspost/compress` for the container (B4) — risk is integration cleanliness with excelize.
7. Chunked read path + read row cache (read-side B1).
8. Admission control + memory budget (§7) — throughput *stability* under load rather than raw speed.

**Tier 3 — everything else, only on benchmark evidence:**
9. Template cache, `sync.Pool` buffers, format-render LRU.
10. Bespoke SAX writer (§3) — only if profiles show StreamWriter dominating.
11. `saveAsync` ticket API.

**Avoided premature optimizations (explicit):** custom allocator/arena for cells, shared-
memory transport instead of CGO marshaling, parallel sheet writing (ZIP is sequential
anyway), JIT-compiled formula engine.

---

## 10. Expected performance vs. the field

> **Phase-0 measurements (2026-06-11, Docker on Apple Silicon, PHP 8.5,
> write N rows × 10 mixed columns, single process):**
>
> | Library | 100k rows | 1M rows | Peak PHP memory |
> |---|---|---|---|
> | PhpSpreadsheet 5.8 | 14.74s | n/a (multi-GB) | 665MB at 100k |
> | rap2hpoutre/fast-excel | 4.00s | — | 4MB |
> | OpenSpout 4.x | 3.64s | 36.74s | 4MB |
> | fast-excel-writer 6.x | 2.67s | 28.16s | 4MB |
> | **easy-excel** | **0.82s** | **7.85s** | **4MB** |
>
> Measured: **3.2–3.6× fast-excel-writer, 4.4–4.7× OpenSpout, ~18× PhpSpreadsheet**
> at OpenSpout-class constant memory — within the estimated bands below
> (1M-row estimate was 8–20s; measured 7.8s). Raw CSV: `bench/results.csv`;
> committed snapshot: `bench/baseline-2026-06-11-php8.5.csv`.

### Original estimates (pre-implementation — kept for accountability)

Reference workload: **write 1,000,000 rows × 10 mixed columns (string/int/float/date), save XLSX**, PHP 8.4 / FrankenPHP worker, M-class or modern x86 server.

| Library | Est. wall time | Est. peak PHP-process memory | Basis |
|---------|---------------|------------------------------|-------|
| PhpSpreadsheet 5.x | ~360–600s | 2–6GB (or cell-cache thrash) | community benchmarks, cell-object model |
| rap2hpoutre/fast-excel (Spout wrapper, + collection overhead) | ~90–150s | ~50–300MB (collections can spike) | wraps OpenSpout; Laravel collection layer adds cost |
| OpenSpout 4.x | ~60–90s | **~3–10MB** (true streaming) | its design target |
| aVadim483/fast-excel-writer 6.x | ~45–75s | ~10–30MB | claims ~1.5–2× OpenSpout; pure-PHP string building |
| **easy-excel (this design, streaming path)** | **~8–20s** | **~10–40MB total (PHP shim buffers + Go StreamWriter)** | excelize StreamWriter benchmarks (102400×50 matrix in seconds); CGO amortized to ~2k calls total; DEFLATE becomes the dominant cost |
| easy-excel, flat ABI w/o shim (power users) | ~6–15s | similar | removes PHP array building in shim |

**Honest summary of the estimate:** ~3–6× faster than fast-excel-writer, ~4–8× faster than
OpenSpout, ~25–50× faster than PhpSpreadsheet, at OpenSpout-class memory — because PHP's
remaining job is just building value arrays, and the XML/ZIP work runs at Go speed. The
floor is set by (a) PHP building 10⁷ zvals for the source data (irreducible, app-side) and
(b) DEFLATE CPU. **Read path** gains are larger still for `toArray` workloads (no PHP cell
objects materialized until requested). These numbers are *estimates with stated mechanisms*;
Phase 0 builds the benchmark rig before any optimization claims are accepted.

Infra cost impact: exports that needed dedicated high-memory queue workers fit in standard
web workers; ~5–10× fewer worker-seconds per export → proportional reduction in
export-fleet cost; constant memory removes the OOM-kill class of incidents.

---

## 11. Repository layout

```
easy-excel/
├── PLAN.md / COMPAT.md / README.md
├── extension/                  # Go module (built via frankenphp extension-init + xcaddy)
│   ├── bridge.go               # //export_php:function layer only
│   ├── registry/               # handle table, request lifecycle
│   ├── core/                   # excelize adapters, dual write paths, style interner
│   ├── compat/                 # PhpSpreadsheet semantic mapping (dates, formats, A1, binders)
│   ├── limits/                 # semaphore, memory budget, timeouts
│   ├── io/                     # path policy, sinks, template cache
│   └── *_test.go               # Go unit tests per package
├── php/                        # Composer package easy-excel/polyfill
│   ├── src/PhpOffice/PhpSpreadsheet/   # shim classes (drop-in namespace via composer replace)
│   ├── src/EasyExcel/          # native API, Hints, Overloaded exception
│   └── tests/                  # PHPUnit: shim tests + imported PhpSpreadsheet suite subset
├── bench/                      # Phase 0 rig: same workloads across all 5 libraries,
│   │                           # docker-compose pinned environment, results committed as CSV
│   └── workloads/              # 1M-row write, 100k read, styled report, template fill
└── .github/workflows/          # build matrix (PHP 8.4/8.5), govulncheck, bench-on-tag
```

## 12. Testing strategy

- **Go unit tests** per package; golden-file XLSX fixtures validated by re-opening with
  excelize *and* by a CI job that opens outputs with real PhpSpreadsheet (round-trip truth).
- **PHPUnit compatibility suite**: import the relevant subset of PhpSpreadsheet's own tests
  (Functional + Cell/Style/Coordinate units), run against the shim. Coverage of that suite
  *is* the compatibility metric, reported in COMPAT.md per release.
- **Concurrency tests**: race detector builds; soak test (10k requests, 64 concurrent) in
  FrankenPHP worker mode asserting zero handle leaks and flat RSS.
- **Security tests**: zip bomb, deep XML, path traversal, stale/forged handle fixtures.
- **Benchmarks as tests**: `bench/` rig runs on every tag; regressions >10% fail the release.

## 13. Execution roadmap

| Phase | Deliverable | Exit criterion |
|-------|------------|----------------|
| **0. Bench rig + walking skeleton** (first) | `extension-init` hello-world built into FrankenPHP; bench harness running all 5 competitor libs on the 4 workloads; baseline CSV committed | reproducible numbers, build pipeline green |
| **1. MVP core** | Tier-1 items: handles, registry+sweep, StreamWriter path, shim with Phase-1 API, Xlsx/Csv write + Xlsx/Csv read | 1M-row workload beats OpenSpout ≥3× at ≤50MB; PhpSpreadsheet Phase-1 test subset green |
| **2. Formatting & load control** | Tier-2: styles+interner, merged cells/autofilter/panes, klauspost zip, admission control, memory budget | styled-report workload parity with PhpSpreadsheet output (binary-diff XML semantics); soak test flat RSS |
| **3. Advanced compat** | formulas via excelize engine + coverage table, charts, validations, template cache, async ticket API (if still justified) | COMPAT.md ≥ agreed coverage; re-measure & publish |

Each phase ends with **re-measurement against the Phase-0 baseline** and a written
what/why/impact note per significant change (response time, throughput, memory, infra cost)
— per the iterative strategy: analyse → plan → implement highest-impact → re-measure → continue.

> **Phase-2 outcome (2026-06-12):** full style graph + structure shipped via a
> per-sheet **op-log** in Go: styles/heights queued before their rows stream
> inline through the StreamWriter (zero degrade for the style-header-then-bulk
> pattern); widths/panes/merges use the StreamWriter's native support; ops
> excelize cannot stream (auto-filter, auto-size, hyperlinks, comments, page
> setup, late styles) defer to save and cost at most one degrade. Divergences
> are listed in COMPAT.md §9–15; the styled-report bench lane is
> `run.php <lib> write-styled`.
>
> Measured (PHP 8.5, Docker/Apple Silicon): styled 100k report — PhpSpreadsheet
> 17.27s / 670MB vs easy-excel **4.30s / 4MB** (4.0×, includes the auto-filter
> degrade). Styled 1M — **12.6s** fully streamed without auto-filter; 121s with
> it (the degrade materializes 1M rows in excelize's random model). Phase-3
> candidate: inject `<autoFilter>` into the streamed sheet XML at save to
> remove that cliff. Inline styling overhead on the hot path: 1M×10 cells with
> a full-column format = 12.6s vs 7.8s unstyled (per-cell style resolution;
> optimizable by pre-filtering entries per batch if profiling justifies).
> **klauspost/compress was deferred to Phase 3**:
> excelize v2.10 exposes no zip-writer hook, so it would mean re-compressing
> the container at save time — and the measured 1M-row write (7.8s, 4.6×
> OpenSpout) shows DEFLATE is not the current bottleneck. Re-evaluate with
> profiling per §3.

> **Phase-3 outcome (2026-06-12):**
>
> - **Streaming auto-filter** shipped: the saved container is patched (raw
>   zip copy + one worksheet rewrite injecting `<autoFilter>` after
>   `</sheetData>`), removing the Phase-2 degrade cliff for filtered streamed
>   exports. The inline styler also gained a per-batch row-uniform fast path
>   (bitmask-cached entry sets instead of per-cell string keys). Measured
>   (same session, machine ~1.5× slower than the Phase-2 session): styled
>   100k 4.30s→1.93s and PhpSpreadsheet ratio 4.0×→**13.9×**; styled 1M
>   121s→**18.5s** (vs 11.9s unstyled control — the filter+styles now cost
>   ~1.6× instead of ~10×).
> - **Formula engine**: bulk calculated reads (`toArray(calculateFormulas:
>   true)`) evaluate uncached formula cells through excelize; coverage table
>   published in FORMULAS.md — **466/529** PhpSpreadsheet functions
>   (generator: `extension/tools/formula-coverage`).
> - **Data validation, conditional formatting (cellIs/containsText/expression
>   + colorScale/dataBar helpers), images with dimension-derived scaling,
>   sheet protection** via PhpSpreadsheet APIs; **charts** via a native
>   declarative API mapped to excelize.AddChart (PhpSpreadsheet's chart
>   object graph deliberately not mapped — COMPAT.md).
> - **Dropped with rationale:** template cache (excelize has no safe model
>   clone; the cache could only hold bytes, which OS page cache already
>   does) and the async ticket API (saves measured at 0.8–13s; FrankenPHP
>   worker + queue patterns cover the rest). Workbook encryption out of
>   scope.

---

## 14. Open questions for approval

1. **Shim namespace strategy**: ship as `composer replace phpoffice/phpspreadsheet` drop-in
   (maximum compat, some legal/naming care needed) vs. `EasyExcel\Compat\…` + class_alias
   bootstrap (safer, one extra install step)? *Recommendation: the latter.*
2. **MVP scope check**: is Phase-1 API list (§5) the right 90% for your workloads, or are
   charts/formulas needed earlier?
3. **Fallback dependency**: is bundling real PhpSpreadsheet as an optional fallback for
   unsupported formats (Ods/Xls) acceptable, or should unsupported formats just throw?
4. Target PHP/FrankenPHP minimum versions (assumed PHP ≥ 8.3, FrankenPHP ≥ 1.9)?
