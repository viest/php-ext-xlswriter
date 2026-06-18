--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("write_boolean.xlsx")
    ->header(['name', 'age', 'active'])
    ->data([
        ['viest', 21, true],
        ['wjx',   21, false],
    ])
    ->output();

var_dump($filePath);

/* Round-trip via the reader: before worksheet_write_boolean() the bool cell
 * fell through type_writer() and was emitted as a blank cell, so it read back
 * as "" instead of 1/0. The reader has no boolean type (true -> int(1),
 * false -> string("0")), so normalise with (int) and explicitly reject the
 * blank-cell regression with !== ''. */
$reader = new \Vtiful\Kernel\Excel($config);
$data   = $reader->openFile("write_boolean.xlsx")->openSheet()->getSheetData();

$active_true  = $data[1][2]; // viest -> true
$active_false = $data[2][2]; // wjx   -> false

var_dump($active_true  !== '' && (int)$active_true  === 1);
var_dump($active_false !== '' && (int)$active_false === 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/write_boolean.xlsx');
?>
--EXPECTF--
string(%d) "%swrite_boolean.xlsx"
bool(true)
bool(true)
