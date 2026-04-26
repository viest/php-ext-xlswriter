# Benchmarks

Lightweight perf harness for php-ext-xlswriter. Each script runs a single
workload and writes a JSON envelope to `bench/results/<git-sha>.json`.

## Run

```sh
# Build the extension first:
phpize && ./configure --enable-reader && make -j

# Default (100k rows × 10 cols ≈ 1M cells):
php -d extension=$(pwd)/modules/xlswriter.so bench/bench_write_1m_cells.php
php -d extension=$(pwd)/modules/xlswriter.so bench/bench_read_1m_cells.php
php -d extension=$(pwd)/modules/xlswriter.so bench/bench_constmemory.php

# Override size (true 1M cells in 27 cols):
BENCH_ROWS=1000000 BENCH_COLS=27 php bench/bench_write_1m_cells.php
```

Override `BENCH_TMP_DIR` to write fixtures somewhere other than
`/tmp`.

## Compare against a baseline

```sh
php bench/compare.php bench/results/<base-sha>.json bench/results/<head-sha>.json
```

Exits 0 always; prints a warning when any benchmark slows down by more
than 15% (override with `BENCH_REGRESSION_THRESHOLD=0.10` etc.). The
intended CI usage is "warn but don't fail" — perf is bursty enough that
hard-failing PRs would be too noisy.

## Adding a benchmark

1. New file `bench/bench_<topic>.php`.
2. `require __DIR__ . '/_lib.php';` and call `bench_require_extension()`.
3. Wrap the hot path in `bench_record('name', function () { ... })`.
4. End with `bench_emit_json(['benchmark' => '<topic>', 'results' => [$r]]);`.
