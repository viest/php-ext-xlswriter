<?php
/* 03 — Streaming read: nextRow() pulls one row at a time, RSS stays flat.
 *
 * The reader returns null when the sheet is exhausted. Each row is a
 * 0-indexed numeric array of cell values.
 */

$dir  = sys_get_temp_dir();
$name = '03_read_streaming.xlsx';

$writer = new \Vtiful\Kernel\Excel(['path' => $dir]);
$writer->fileName($name)->header(['n', 'square']);
for ($i = 1; $i <= 1000; $i++) {
    $writer->data([[$i, $i * $i]]);
}
$writer->output();

$reader = new \Vtiful\Kernel\Excel(['path' => $dir]);
$reader->openFile($name)->openSheet();

$count = 0;
$last  = null;
while (($row = $reader->nextRow()) !== null) {
    $last = $row;
    $count++;
}
printf("read %d rows; last = [%s]\n", $count, implode(', ', $last));

@unlink($dir . '/' . $name);
