<?php
/* 07 — CSV export: stream a sheet directly to a stdio file pointer.
 *
 *   putCSV($fp [, $delim, $enclosure, $escape])
 *   putCSVCallback(callable $cb, $fp [, $delim, $enclosure, $escape])
 *
 * The callback variant lets you mutate each row before it gets written
 * (e.g. mask columns, compute derived fields).
 */

$dir  = sys_get_temp_dir();
$name = '07_csv_export.xlsx';
$csv  = $dir . '/07_csv_export.csv';

$writer = new \Vtiful\Kernel\Excel(['path' => $dir]);
$writer->fileName($name)
    ->header(['name', 'amount', 'currency'])
    ->data([
        ['Alice', 1234.56, 'USD'],
        ['Bob',     987.65, 'EUR'],
    ])
    ->output();

$reader = new \Vtiful\Kernel\Excel(['path' => $dir]);
$reader->openFile($name)->openSheet();

$fp = fopen($csv, 'w');
$reader->putCSV($fp);
fclose($fp);

echo file_get_contents($csv);

@unlink($csv);
@unlink($dir . '/' . $name);
