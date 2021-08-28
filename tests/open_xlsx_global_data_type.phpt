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
$filePath = $excel->fileName('open_xlsx_global_data_type.xlsx')
    ->header(['Item', 'Cost', 'Int', 'Double'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->output();

$excel->openFile('open_xlsx_global_data_type.xlsx')
    ->setGlobalType(\Vtiful\Kernel\Excel::TYPE_DOUBLE)
    ->nextCellCallback(function ($row, $cell, $data) {
        echo 'cell:' . $cell . ', row:' . $row .', data type:' . gettype($data) . PHP_EOL;
    });

echo '----------------' . PHP_EOL;

$data = $excel->openFile('open_xlsx_global_data_type.xlsx')
    ->openSheet()
    ->setGlobalType(\Vtiful\Kernel\Excel::TYPE_STRING)
    ->getSheetData();

var_dump($data);

$excel->openFile('open_xlsx_global_data_type.xlsx')
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
@unlink(__DIR__ . '/open_xlsx_global_data_type.xlsx');
?>
--EXPECT--
cell:0, row:0, data type:string
cell:1, row:0, data type:string
cell:2, row:0, data type:string
cell:3, row:0, data type:string
cell:3, row:0, data type:string
cell:0, row:1, data type:string
cell:1, row:1, data type:string
cell:2, row:1, data type:double
cell:3, row:1, data type:double
cell:3, row:1, data type:string
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