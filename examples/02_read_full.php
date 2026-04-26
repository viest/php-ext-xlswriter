<?php
/* 02 — Full read: getSheetData() loads the whole sheet into memory.
 *
 * Use this for small/medium files. For anything beyond a few tens of
 * thousands of rows prefer the streaming pattern in 03_read_streaming.php.
 */

$dir  = sys_get_temp_dir();
$name = '02_read_full.xlsx';

$writer = new \Vtiful\Kernel\Excel(['path' => $dir]);
$writer->fileName($name)
    ->header(['col_a', 'col_b'])
    ->data([['x', 1], ['y', 2], ['z', 3]])
    ->output();

$reader = new \Vtiful\Kernel\Excel(['path' => $dir]);
$rows   = $reader->openFile($name)->openSheet()->getSheetData();
print_r($rows);

@unlink($dir . '/' . $name);
