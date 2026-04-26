<?php
/*
 * Bench: streaming read of a large grid via nextRow().
 *
 * Generates the fixture itself if not present. Defaults to 100k × 10.
 * Override via env: BENCH_ROWS / BENCH_COLS.
 */
require __DIR__ . '/_lib.php';
bench_require_extension();

if (!method_exists(\Vtiful\Kernel\Excel::class, 'openFile')) {
    fwrite(STDERR, "reader not enabled in this build (--disable-reader)\n");
    exit(0);
}

$rows = (int) (getenv('BENCH_ROWS') ?: 100000);
$cols = (int) (getenv('BENCH_COLS') ?: 10);
$dir  = bench_tmp_dir();
$path = $dir . '/bench_read.xlsx';

if (!is_file($path)) {
    $excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
    $excel->fileName('bench_read.xlsx');
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
    unset($excel);
}

$result = bench_record('read_stream', function () use ($dir, $path) {
    $rd = new \Vtiful\Kernel\Excel(['path' => $dir]);
    $rd->openFile('bench_read.xlsx')->openSheet();
    $count = 0;
    while (($row = $rd->nextRow()) !== null) {
        $count++;
    }
    return ['rows_read' => $count, 'file_bytes' => filesize($path)];
});

@unlink($path);

bench_emit_json([
    'benchmark' => 'read_1m_cells',
    'rows'      => $rows,
    'cols'      => $cols,
    'results'   => [$result],
]);
