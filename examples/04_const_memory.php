<?php
/* 04 — Constant-memory write: trades CPU for tiny RSS. Use when writing
 *      very large files (1M+ rows) that would otherwise blow PHP's
 *      memory_limit.
 *
 * The trade-off vs. normal mode (29s + 2 GB) is roughly 50s + <1 MB on
 * a 1M-rows × 27-cols × 19-byte-string workload (see README benchmark).
 */

$dir  = sys_get_temp_dir();
$name = '04_const_memory.xlsx';

$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel = $excel->constMemory($name);

$excel->header(['id', 'value']);

$rows = (int) (getenv('EXAMPLE_ROWS') ?: 50000);
for ($i = 1; $i <= $rows; $i++) {
    $excel->data([[$i, "row-$i"]]);
}
$path = $excel->output();

printf("wrote %d rows; file size = %d bytes; peak php mem = %d KiB\n",
    $rows, filesize($path), memory_get_peak_usage(true) / 1024);

@unlink($path);
