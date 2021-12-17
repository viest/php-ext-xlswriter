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
$filePath = $excel->fileName('open_xlsx_get_data_bignumbers.xlsx')
    ->header(['Item'])
    ->data([
        ['9999999999999999999999999999999999999999999999999999999999999'],
    ])
    ->output();

$data = $excel->openFile('open_xlsx_get_data_bignumbers.xlsx')
    ->openSheet()
    ->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_data_bignumbers.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(4) "Item"
  }
  [1]=>
  array(1) {
    [0]=>
    string(61) "9999999999999999999999999999999999999999999999999999999999999"
  }
}
