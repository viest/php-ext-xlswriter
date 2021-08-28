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
$filePath = $excel->fileName('open_xlsx_get_sheet_not_found_data.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

$data = $excel->openFile('open_xlsx_get_sheet_not_found_data.xlsx')
    ->openSheet('not_found')->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_sheet_not_found_data.xlsx');
?>
--EXPECT--
array(0) {
}