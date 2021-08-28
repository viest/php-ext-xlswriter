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
$filePath = $excel->fileName('xlsx_to_csv.xlsx', 'TestSheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->output();

$fp = fopen('./tests/file.csv', 'w');

$csvResult = $excel->openFile('xlsx_to_csv.xlsx')
    ->openSheet()
    ->putCSV($fp);

var_dump($csvResult);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/xlsx_to_csv.xlsx');
@unlink(__DIR__ . '/file.csv');
?>
--EXPECT--
bool(true)
