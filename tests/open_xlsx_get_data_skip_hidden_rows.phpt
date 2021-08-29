--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests/xlsx'];
$excel  = new \Vtiful\Kernel\Excel($config);

$data = $excel->openFile('hidden_row.xlsx')
    ->openSheet('Sheet1')
    ->getSheetData();

var_dump($data);

$data = $excel->openFile('hidden_row.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_HIDDEN_ROW|\Vtiful\Kernel\Excel::SKIP_EMPTY_ROW)
    ->getSheetData();

var_dump($data);
?>
--EXPECT--
array(4) {
  [0]=>
  array(1) {
    [0]=>
    string(4) "name"
  }
  [1]=>
  array(1) {
    [0]=>
    string(8) "ZhangSan"
  }
  [2]=>
  array(1) {
    [0]=>
    string(4) "LiSi"
  }
  [3]=>
  array(1) {
    [0]=>
    string(6) "WangWu"
  }
}
array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(4) "name"
  }
  [1]=>
  array(1) {
    [0]=>
    string(6) "WangWu"
  }
}
