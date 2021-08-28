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
$filePath = $excel->fileName('open_xlsx_get_data_skip_empty.xlsx')
    ->header(['', 'Cost'])
    ->data([
        [],
        ['viest', '']
    ])
    ->output();

$data = $excel->openFile('open_xlsx_get_data_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS)
    ->getSheetData();

var_dump($data);

$data = $excel->openFile('open_xlsx_get_data_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW)
    ->getSheetData();

var_dump($data);

$data = $excel->openFile('open_xlsx_get_data_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS|\Vtiful\Kernel\Excel::SKIP_EMPTY_ROW)
    ->getSheetData();

var_dump($data);

$data = $excel->openFile('open_xlsx_get_data_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_VALUE)
    ->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_data_skip_empty.xlsx');
?>
--EXPECT--
array(3) {
  [0]=>
  array(1) {
    [1]=>
    string(4) "Cost"
  }
  [1]=>
  array(0) {
  }
  [2]=>
  array(1) {
    [0]=>
    string(5) "viest"
  }
}
array(2) {
  [0]=>
  array(2) {
    [0]=>
    string(0) ""
    [1]=>
    string(4) "Cost"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    string(0) ""
  }
}
array(2) {
  [0]=>
  array(1) {
    [1]=>
    string(4) "Cost"
  }
  [1]=>
  array(1) {
    [0]=>
    string(5) "viest"
  }
}
array(3) {
  [0]=>
  array(1) {
    [1]=>
    string(4) "Cost"
  }
  [1]=>
  array(0) {
  }
  [2]=>
  array(1) {
    [0]=>
    string(5) "viest"
  }
}