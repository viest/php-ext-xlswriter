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
$filePath = $excel->fileName('open_xlsx_get_data_with_set_type.xlsx')
    ->header(['Name', 'Age', 'Date'])
    ->data([
        ['Viest', 24]
    ])
    ->insertDate(1, 2, 1568877706)
    ->output();

$data = $excel->openFile('open_xlsx_get_data_with_set_type.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
    ])
    ->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_data_with_set_type.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  array(3) {
    [0]=>
    string(4) "Name"
    [1]=>
    string(3) "Age"
    [2]=>
    string(4) "Date"
  }
  [1]=>
  array(3) {
    [0]=>
    string(5) "Viest"
    [1]=>
    string(2) "24"
    [2]=>
    int(1568877706)
  }
}