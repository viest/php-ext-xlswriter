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
$filePath = $excel->fileName('open_xlsx_get_data_skip_rows.xlsx', 'TestSheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
        ['Item_2', 'Cost_2', 10, 10.9999995],
        ['Item_3', 'Cost_3', 10, 10.9999995],
    ])
    ->output();

$data = $excel->openFile('open_xlsx_get_data_skip_rows.xlsx')
    ->openSheet()
    ->setSkipRows(3)
    ->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_data_skip_rows.xlsx');
?>
--EXPECT--
array(1) {
  [0]=>
  array(4) {
    [0]=>
    string(6) "Item_3"
    [1]=>
    string(6) "Cost_3"
    [2]=>
    int(10)
    [3]=>
    float(10.9999995)
  }
}