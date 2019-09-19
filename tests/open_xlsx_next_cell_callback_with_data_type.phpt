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
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost', 'Int', 'Double', 'Date'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->insertDate(1, 4, 1568904314)
    ->output();

$excel->openFile('tutorial.xlsx')
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_INT,
        \Vtiful\Kernel\Excel::TYPE_DOUBLE,
        \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
    ])
    ->nextCellCallback(function ($row, $cell, $data) {
        echo 'cell:' . $cell . ', row:' . $row . PHP_EOL;
        var_dump($data);
    });
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
cell:0, row:0
string(4) "Item"
cell:1, row:0
string(4) "Cost"
cell:2, row:0
string(3) "Int"
cell:3, row:0
string(6) "Double"
cell:4, row:0
string(4) "Date"
cell:4, row:0
string(12) "XLSX_ROW_END"
cell:0, row:1
string(6) "Item_1"
cell:1, row:1
string(6) "Cost_1"
cell:2, row:1
int(10)
cell:3, row:1
float(10.9999995)
cell:4, row:1
int(1568904314)
cell:4, row:1
string(12) "XLSX_ROW_END"