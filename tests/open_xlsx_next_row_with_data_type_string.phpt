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
$filePath = $excel->fileName('open_xlsx_next_row_with_data_type_string.xlsx')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1'],
    ])
    ->output();

$excel->openFile('open_xlsx_next_row_with_data_type_string.xlsx')
    ->openSheet();

var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
var_dump($excel->nextRow([\Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING]));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row_with_data_type_string.xlsx');
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