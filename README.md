# easy-excel

A PHP spreadsheet extension written in **Go** for [FrankenPHP](https://frankenphp.dev),
exposing a **PhpSpreadsheet-compatible API** while doing all heavy lifting (XML
generation, ZIP compression, parsing, formulas) in compiled Go via
[excelize](https://github.com/qax-os/excelize).

Target: PhpSpreadsheet's ergonomics at OpenSpout-class (constant) memory and a
multiple of both libraries' throughput. Design rationale, bottleneck analysis and
measured claims live in [PLAN.md](PLAN.md); API coverage in [COMPAT.md](COMPAT.md).

## How it works

```
PhpOffice\PhpSpreadsheet\* (aliases)        ← your existing code, unchanged
        │
EasyExcel\Compat\* PHP shim                 ← tiny objects, write-behind row buffer
        │  flat easy_excel_*() calls, batched rows (1 CGO call / 512 rows)
Go extension: handle registry → excelize    ← StreamWriter, style interner,
                                              admission control, path policy
```

Key properties:

- **Dual write path** — sequential writes stream at constant memory; an
  out-of-order write triggers a one-time documented fallback to random access.
- **Load control** — heavy operations pass a weighted semaphore + memory
  budget; overload surfaces as `EasyExcel\Exception\Overloaded` (map it to
  HTTP 429 or your queue) instead of OOM-killing the worker.
- **Security defaults** — unzip-bomb limits, path allowlist
  (`EASY_EXCEL_ALLOWED_PATHS`), capability-style random handles, opt-in CSV
  injection guard.

## Quick start (Docker)

```bash
# run the test suites
make test                 # = docker build --target=go-test / --target=php-test

# build FrankenPHP with the extension baked in
make build                # produces the frankenphp-easy-excel image

# or pull the published image
docker pull ghcr.io/xiidea/frankenphp8.5-easy-excel:latest
```

Use the shim in your app (the image ships it at `/opt/easy-excel/php`):

```json
{
    "repositories": [{ "type": "path", "url": "/opt/easy-excel/php" }],
    "require": { "easy-excel/polyfill": "*" }
}
```

```php
use PhpOffice\PhpSpreadsheet\Spreadsheet;          // resolved to EasyExcel\Compat\*
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->fromArray($hugeDataset);                    // batched straight into Go
$sheet->setCellValue('A1', 'Hello');                // buffered, flushed in row batches
IOFactory::createWriter($spreadsheet, 'Xlsx')->save('report.xlsx');
$spreadsheet->disconnectWorksheets();               // frees the native workbook
```

The aliases stay dormant when the real `phpoffice/phpspreadsheet` is installed
or the extension is missing, so adoption can be incremental.

## Configuration

| Env var | Default | Meaning |
|---|---|---|
| `EASY_EXCEL_MAX_CONCURRENT` | `max(2, NumCPU)` | heavy ops (open/save/scan) in flight |
| `EASY_EXCEL_ACQUIRE_TIMEOUT_MS` | `30000` | wait before `Overloaded` is raised |
| `EASY_EXCEL_MEMORY_BUDGET_MB` | `512` | estimated live-workbook bytes circuit breaker |
| `EASY_EXCEL_ALLOWED_PATHS` | unset (any local path) | colon-separated base dirs for load/save |

## Repository layout

```
extension/   Go module: registry, limits, exio, compat, core + easy_excel.go (bridge)
php/         Composer package: EasyExcel\Compat shim + alias bootstrap + tests
bench/       Phase-0 rig: identical workloads across 5 libraries
Dockerfile   go-test | php-test | generate | build | runner stages
```

## Development

```bash
make test            # Docker: Go (race) + PHP suites
make host-test       # local toolchains, faster iteration
make bench           # baseline all libraries (results.csv)
```

Notes:

- `extension/easy_excel.go` is **excluded from gofmt**: gofmt mangles
  `//export_php:` directives (underscores aren't recognized as directive
  names). `make fmt` formats only the pure packages.
- The bridge compiles only inside the `generate`/`build` Docker stages (needs
  PHP ZTS headers); the five pure-Go packages build and test anywhere.
- CI publishes `ghcr.io/<owner>/frankenphp-easy-excel` on main pushes and
  semver tags (see `.github/workflows/publish.yml`).

## Measured performance

Write N rows × 10 mixed columns, one process per run (Docker, PHP 8.5,
Apple Silicon; `bench/baseline-2026-06-11-php8.5.csv`):

| Library | 100k rows | 1M rows | Peak PHP memory |
|---|---|---|---|
| PhpSpreadsheet 5.8 | 14.74s | — | 665MB at 100k |
| rap2hpoutre/fast-excel | 4.00s | — | 4MB |
| OpenSpout 4.x | 3.64s | 36.74s | 4MB |
| fast-excel-writer 6.x | 2.67s | 28.16s | 4MB |
| **easy-excel** | **0.82s** | **7.85s** | **4MB** |

Styled report (same data + styled header, two column number formats, widths,
freeze pane, auto-filter; `write-styled` workload,
`bench/baseline-2026-06-12-php8.5-phase3.csv`; this session ran ~1.5× slower
than the table above — its unstyled 1M control was 11.9s — so compare ratios,
not sessions):

| Library | 100k rows | 1M rows | Peak PHP memory |
|---|---|---|---|
| PhpSpreadsheet 5.8 | 26.79s | — | 670MB at 100k |
| **easy-excel** | **1.93s** (13.9×) | **18.5s** | **4MB** |

Styles stream inline, and since Phase 3 the auto-filter no longer degrades
either: the `<autoFilter>` element is injected into the saved container
directly (COMPAT.md §16). In Phase 2 the same styled 1M lane took 121s; the
streaming filter and the batch-uniform style resolver cut it to ~1.6× the
unstyled write.

## Status

Phase 3 (advanced compat) complete, on top of Phase 2's full style graph and
worksheet structure:

- **Formulas**: `getCalculatedValue()` and bulk `toArray(calculateFormulas:
  true)` via excelize's engine — **466/529** PhpSpreadsheet functions, per-
  function table in [FORMULAS.md](FORMULAS.md).
- **Data validation, conditional formatting, images, sheet protection** with
  the PhpSpreadsheet APIs; **charts** via a native declarative API
  (`addNativeChart`).
- **Streaming auto-filter**: filtered million-row exports no longer degrade —
  the filter is injected into the saved container.

Styles applied before their rows are written are inlined into the stream, so
the usual "style header → bulk write" report keeps constant memory and full
write speed. See COMPAT.md for the precise matrix and documented divergences.
Reproduce the numbers with `make bench`.
