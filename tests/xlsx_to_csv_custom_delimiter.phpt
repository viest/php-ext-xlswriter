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
$filePath = $excel->fileName('xlsx_to_csv_custom_delimiter.xlsx', 'TestSheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->output();

$fp = fopen('./tests/file.csv', 'w');

$csvResult = $excel->openFile('xlsx_to_csv_custom_delimiter.xlsx')
    ->openSheet()
    ->putCSV($fp, ';');

var_dump($csvResult);

if (($csvHandler = fopen('./tests/file.csv', 'r')) === FALSE) {
    die('csv file open failure');
}

while (($data = fgetcsv($csvHandler, 1000, ';')) !== FALSE) {
    var_dump($data);
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/xlsx_to_csv_custom_delimiter.xlsx');
@unlink(__DIR__ . '/file.csv');
?>
--EXPECT--
bool(true)
array(2) {
  [0]=>
  string(4) "Item"
  [1]=>
  string(4) "Cost"
}
array(4) {
  [0]=>
  string(6) "Item_1"
  [1]=>
  string(6) "Cost_1"
  [2]=>
  string(2) "10"
  [3]=>
  string(10) "10.9999995"
}
