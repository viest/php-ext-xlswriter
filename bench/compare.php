<?php
/*
 * Compare two benchmark JSON files. Exits 0 always; prints a warning
 * when any benchmark regresses by more than the threshold (default 15%).
 *
 * Usage:
 *   php bench/compare.php bench/results/<base>.json bench/results/<head>.json
 */
if ($argc < 3) {
    fwrite(STDERR, "usage: php bench/compare.php <base.json> <head.json>\n");
    exit(2);
}

$threshold = (float) (getenv('BENCH_REGRESSION_THRESHOLD') ?: 0.15);

$base = json_decode((string) file_get_contents($argv[1]), true);
$head = json_decode((string) file_get_contents($argv[2]), true);
if (!$base || !$head) {
    fwrite(STDERR, "failed to parse JSON\n");
    exit(2);
}

$index = static function (array $env): array {
    $out = [];
    foreach ($env['benchmarks'] ?? [] as $bench_name => $bench) {
        foreach ($bench['results'] ?? [] as $r) {
            $key       = $bench_name . '/' . ($r['name'] ?? 'unnamed');
            $out[$key] = $r;
        }
    }
    return $out;
};

$b = $index($base);
$h = $index($head);

$warnings = 0;
foreach ($h as $name => $row) {
    if (!isset($b[$name])) {
        continue;
    }
    $bs = $b[$name]['secs'];
    $hs = $row['secs'];
    if ($bs <= 0) continue;
    $delta = ($hs - $bs) / $bs;
    $arrow = $delta >= 0 ? '+' : '';
    printf("%-24s base=%.3fs head=%.3fs (%s%.1f%%)\n",
        $name, $bs, $hs, $arrow, $delta * 100);
    if ($delta > $threshold) {
        $warnings++;
        fprintf(STDERR, "  WARN: %s regressed by %.1f%% (threshold %.0f%%)\n",
            $name, $delta * 100, $threshold * 100);
    }
}

if ($warnings > 0) {
    fwrite(STDERR, "$warnings benchmark(s) regressed beyond threshold\n");
}
exit(0);
