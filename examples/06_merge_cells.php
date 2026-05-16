<?php
/* 06 — Merge cells: pass an A1 range and the master cell value. */

$dir  = sys_get_temp_dir();
$name = '06_merge_cells.xlsx';

$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel->fileName($name)
    ->mergeCells('A1:D1', 'Q1 2026 Report')
    ->header(['Region', 'Sales', 'Returns', 'Net'])
    ->data([
        ['North', 12500, 320, 12180],
        ['South',  9800, 110,  9690],
    ])
    ->output();

echo "wrote $dir/$name\n";
@unlink($dir . '/' . $name);
