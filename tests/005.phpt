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
$filePath = $excel->fileName('005.xlsx')
    ->header(['Item', 'Cost'])
    ->output();
var_dump($filePath);

/* Round-trip: header row reads back exactly as written. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('005.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/005.xlsx');
?>
--EXPECT--
string(16) "./tests/005.xlsx"
array(1) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "Item"
    [1]=>
    string(4) "Cost"
  }
}
