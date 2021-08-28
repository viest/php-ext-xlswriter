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
$filePath = $excel->fileName('open_xlsx_sheet_flag.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

$data = $excel->openFile('open_xlsx_sheet_flag.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS);

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_sheet_flag.xlsx');
?>
--EXPECT--
object(Vtiful\Kernel\Excel)#1 (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(33) "./tests/open_xlsx_sheet_flag.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}