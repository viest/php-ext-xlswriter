<?php
/*
 * Bench: constant-memory write mode.
 *
 * The const-memory mode trades CPU for tiny RSS. We measure both axes so a
 * regression in either direction shows up in the JSON.
 */
require __DIR__ . '/_lib.php';
bench_require_extension();

$rows = (int) (getenv('BENCH_ROWS') ?: 100000);
$cols = (int) (getenv('BENCH_COLS') ?: 10);
$dir  = bench_tmp_dir();
$path = $dir . '/bench_constmem.xlsx';

$result = bench_record('write_constmem', function () use ($rows, $cols, $dir, $path) {
    $excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
    $excel = $excel->constMemory('bench_constmem.xlsx');

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
}, ['mode' => 'const_memory']);

@unlink($path);

bench_emit_json([
    'benchmark' => 'constmemory',
    'results'   => [$result],
]);
