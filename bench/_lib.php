<?php
/* Shared helpers for the benchmark scripts. */

function bench_require_extension(): void
{
    if (!extension_loaded('xlswriter')) {
        fwrite(STDERR, "xlswriter extension is not loaded\n");
        exit(1);
    }
}

function bench_tmp_dir(): string
{
    $dir = getenv('BENCH_TMP_DIR') ?: sys_get_temp_dir();
    if (!is_dir($dir)) {
        mkdir($dir, 0700, true);
    }
    return rtrim($dir, '/');
}

function bench_now(): float
{
    return microtime(true);
}

/**
 * Run the supplied callable, return a result envelope:
 *   {name, rows, cols, secs, mem_peak_bytes, file_bytes?}
 */
function bench_record(string $name, callable $fn, array $extra = []): array
{
    gc_collect_cycles();
    $mem0 = memory_get_peak_usage(true);
    $t0   = bench_now();
    $extra_returned = $fn();
    $t1   = bench_now();
    $mem1 = memory_get_peak_usage(true);

    return array_merge([
        'name'            => $name,
        'secs'            => round($t1 - $t0, 4),
        'mem_peak_bytes'  => $mem1,
        'mem_delta_bytes' => $mem1 - $mem0,
    ], $extra, is_array($extra_returned) ? $extra_returned : []);
}

function bench_emit_json(array $report): void
{
    $sha = trim((string) @shell_exec('git rev-parse --short HEAD 2>/dev/null')) ?: 'unknown';
    $out_dir = __DIR__ . '/results';
    if (!is_dir($out_dir)) {
        mkdir($out_dir, 0755, true);
    }
    $out = $out_dir . '/' . $sha . '.json';

    /* Merge with any previous bench run for the same SHA — multiple bench
     * scripts append into the same file under "benchmarks": {<name>: ...}. */
    $envelope = [
        'git_sha'   => $sha,
        'php'       => PHP_VERSION,
        'os'        => PHP_OS,
        'arch'      => php_uname('m'),
        'benchmarks' => [],
    ];
    if (is_file($out)) {
        $prev = json_decode((string) file_get_contents($out), true);
        if (is_array($prev) && isset($prev['benchmarks']) && is_array($prev['benchmarks'])) {
            $envelope = array_merge($envelope, $prev);
        }
    }

    $name = $report['benchmark'] ?? 'unnamed';
    $report['timestamp'] = date(DATE_ATOM);
    $envelope['benchmarks'][$name] = $report;
    $envelope['updated_at']        = date(DATE_ATOM);

    file_put_contents($out, json_encode($envelope, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES));
    echo "wrote $out (benchmark=$name)\n";
}
