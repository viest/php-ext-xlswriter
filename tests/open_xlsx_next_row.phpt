--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('open_xlsx_next_row.xlsx')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1'],
    ])
    ->output();

$excel->openFile('open_xlsx_next_row.xlsx')
    ->openSheet();

var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  string(4) "Item"
  [1]=>
  string(4) "Cost"
}
array(2) {
  [0]=>
  string(6) "Item_1"
  [1]=>
  string(6) "Cost_1"
}
NULL
NULL
NULL
NULL
NULL