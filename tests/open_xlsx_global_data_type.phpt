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
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost', 'Int', 'Double'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->output();

$excel->openFile('tutorial.xlsx')
    ->setGlobalType(\Vtiful\Kernel\Excel::TYPE_DOUBLE)
    ->nextCellCallback(function ($row, $cell, $data) {
        echo 'cell:' . $cell . ', row:' . $row . PHP_EOL;
        var_dump($data);
    });

echo '----------------' . PHP_EOL;

$data = $excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->setGlobalType(\Vtiful\Kernel\Excel::TYPE_STRING)
    ->getSheetData();

var_dump($data);

$excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->setGlobalType(\Vtiful\Kernel\Excel::TYPE_INT);

echo '----------------' . PHP_EOL;

var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
cell:0, row:0
string(4) "Item"
cell:1, row:0
string(4) "Cost"
cell:2, row:0
string(3) "Int"
cell:3, row:0
string(6) "Double"
cell:3, row:0
string(12) "XLSX_ROW_END"
cell:0, row:1
string(6) "Item_1"
cell:1, row:1
string(6) "Cost_1"
cell:2, row:1
float(10)
cell:3, row:1
float(10.9999995)
cell:3, row:1
string(12) "XLSX_ROW_END"
----------------
array(2) {
  [0]=>
  array(4) {
    [0]=>
    string(4) "Item"
    [1]=>
    string(4) "Cost"
    [2]=>
    string(3) "Int"
    [3]=>
    string(6) "Double"
  }
  [1]=>
  array(4) {
    [0]=>
    string(6) "Item_1"
    [1]=>
    string(6) "Cost_1"
    [2]=>
    string(2) "10"
    [3]=>
    string(10) "10.9999995"
  }
}
----------------
array(4) {
  [0]=>
  string(4) "Item"
  [1]=>
  string(4) "Cost"
  [2]=>
  string(3) "Int"
  [3]=>
  string(6) "Double"
}
array(4) {
  [0]=>
  string(6) "Item_1"
  [1]=>
  string(6) "Cost_1"
  [2]=>
  int(10)
  [3]=>
  int(10)
}
NULL
NULL