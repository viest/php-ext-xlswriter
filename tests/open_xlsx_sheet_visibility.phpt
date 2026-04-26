--TEST--
sheetListWithMeta returns visibility state for each sheet
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$excel->fileName('open_xlsx_sheet_visibility.xlsx', 'first_visible')
    ->addSheet('hidden_sheet')
    ->setCurrentSheetHide()
    ->addSheet('last_visible')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_sheet_visibility.xlsx');

var_dump($reader->sheetList());
var_dump($reader->sheetListWithMeta());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_sheet_visibility.xlsx');
?>
--EXPECT--
array(3) {
  [0]=>
  string(13) "first_visible"
  [1]=>
  string(12) "hidden_sheet"
  [2]=>
  string(12) "last_visible"
}
array(3) {
  [0]=>
  array(2) {
    ["name"]=>
    string(13) "first_visible"
    ["state"]=>
    string(7) "visible"
  }
  [1]=>
  array(2) {
    ["name"]=>
    string(12) "hidden_sheet"
    ["state"]=>
    string(6) "hidden"
  }
  [2]=>
  array(2) {
    ["name"]=>
    string(12) "last_visible"
    ["state"]=>
    string(7) "visible"
  }
}
