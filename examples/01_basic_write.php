<?php
/* 01 — Basic write: header + data table → output() returns the path. */

$dir  = sys_get_temp_dir();
$name = '01_basic_write.xlsx';

$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);

$path = $excel->fileName($name)
    ->header(['Name', 'Email', 'Score'])
    ->data([
        ['Alice', 'alice@example.com', 92],
        ['Bob',   'bob@example.com',   78],
        ['Carol', 'carol@example.com', 85],
    ])
    ->output();

echo "wrote $path\n";
@unlink($path);
