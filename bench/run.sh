#!/usr/bin/env bash
# Phase-0 baseline: run every library over the write workload (and read where
# supported), one process per run, appending to results.csv.
#
#   ./run.sh [rows...]            # default: 10000 100000 1000000
#
# easy-excel rows require a PHP built with the extension; set EASY_EXCEL_PHP
# to that binary (e.g. a frankenphp build), otherwise those runs are skipped.
set -euo pipefail
cd "$(dirname "$0")"

ROWS=("${@:-10000 100000 1000000}")
[ $# -gt 0 ] && ROWS=("$@")

PHP_BIN="${PHP_BIN:-php}"
OUT="results.csv"

[ -d vendor ] || composer install --quiet

echo "lib,workload,rows,wall_seconds,peak_mem_bytes,file_bytes" | tee "$OUT"

for rows in ${ROWS[@]}; do
    for lib in phpspreadsheet openspout fast-excel-writer rap2hpoutre; do
        # PhpSpreadsheet at 1M rows can exhaust memory; cap it explicitly
        "$PHP_BIN" -d memory_limit=6G run.php "$lib" write "$rows" | tee -a "$OUT" \
            || echo "$lib,write,$rows,FAILED,," | tee -a "$OUT"
    done
    if [ -n "${EASY_EXCEL_PHP:-}" ]; then
        # may contain a subcommand, e.g. "frankenphp php-cli" — keep unquoted
        $EASY_EXCEL_PHP run.php easy-excel write "$rows" | tee -a "$OUT" \
            || echo "easy-excel,write,$rows,FAILED,," | tee -a "$OUT"
    else
        echo "# easy-excel skipped: set EASY_EXCEL_PHP to a PHP/FrankenPHP binary with the extension" >&2
    fi
done
