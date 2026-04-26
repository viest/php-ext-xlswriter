<?php
/*
 * Bench: write a large grid in normal (in-memory) mode.
 *
 * Default 100k rows × 10 cols (1M cells). Override via env:
 *   BENCH_ROWS=1000000 BENCH_COLS=27 php bench/bench_write_1m_cells.php
 */
require __DIR__ . '/_lib.php';
bench_require_extension();

$rows = (int) (getenv('BENCH_ROWS') ?: 100000);
$cols = (int) (getenv('BENCH_COLS') ?: 10);
$dir  = bench_tmp_dir();
$path = $dir . '/bench_write.xlsx';

$result = bench_record('write_normal', function () use ($rows, $cols, $dir, $path) {
    $excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
    $excel->fileName('bench_write.xlsx');

    $header = [];
    for ($c = 0; $c < $cols; $c++) {
        $header[] = "col$c";
    }
    $excel->header($header);

    for ($r = 0; $r < $rows; $r++) {
        $row = [];
        for ($c = 0; $c < $cols; $c++) {
            $row[] = $r * $cols + $c;
        }
        $excel->data([$row]);
    }
    $excel->output();

    return ['rows' => $rows, 'cols' => $cols, 'file_bytes' => filesize($path)];
}, ['mode' => 'normal']);

@unlink($path);

bench_emit_json([
    'benchmark' => 'write_1m_cells',
    'results'   => [$result],
]);
