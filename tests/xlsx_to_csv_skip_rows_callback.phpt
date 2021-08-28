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
$filePath = $excel->fileName('xlsx_to_csv_skip_rows_callback.xlsx', 'TestSheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
        ['Item_2', 'Cost_2', 10, 10.9999995],
        ['Item_3', 'Cost_3', 10, 10.9999995],
    ])
    ->output();

$fp = fopen('./tests/file.csv', 'w');

$csvResult = $excel->openFile('xlsx_to_csv_skip_rows_callback.xlsx')
    ->openSheet()
    ->setSkipRows(3)
    ->putCSVCallback(function($row){
        return $row;
    }, $fp);

fclose($fp);

var_dump($csvResult);
var_dump(file_get_contents('./tests/file.csv'));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/xlsx_to_csv_skip_rows_callback.xlsx');
@unlink(__DIR__ . '/file.csv');
?>
--EXPECT--
bool(true)
string(28) "Item_3,Cost_3,10,10.9999995
"