--TEST--
getMergedCells round-trip: writer mergeCells -> reader returns ranges
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$excel->fileName('open_xlsx_get_merged_cells.xlsx', 'Sheet1')
    ->mergeCells('A1:C1', 'header')
    ->mergeCells('D2:E5', 'block')
    ->mergeCells('B7:B9', 'col')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_merged_cells.xlsx')
    ->openSheet();

var_dump($reader->getMergedCells());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_merged_cells.xlsx');
?>
--EXPECT--
array(3) {
  [0]=>
  array(4) {
    ["first_row"]=>
    int(1)
    ["first_col"]=>
    int(1)
    ["last_row"]=>
    int(1)
    ["last_col"]=>
    int(3)
  }
  [1]=>
  array(4) {
    ["first_row"]=>
    int(2)
    ["first_col"]=>
    int(4)
    ["last_row"]=>
    int(5)
    ["last_col"]=>
    int(5)
  }
  [2]=>
  array(4) {
    ["first_row"]=>
    int(7)
    ["first_col"]=>
    int(2)
    ["last_row"]=>
    int(9)
    ["last_col"]=>
    int(2)
  }
}
