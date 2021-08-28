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
$filePath = $excel->fileName('open_xlsx_next_cell_callback_with_data_type.xlsx')
    ->header(['Item', 'Cost', 'Int', 'Double', 'Date'])
    ->data([
        ['Item_1', 'Cost_1', 10, 10.9999995],
    ])
    ->insertDate(1, 4, 1568904314)
    ->output();

$excel->openFile('open_xlsx_next_cell_callback_with_data_type.xlsx')
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_INT,
        \Vtiful\Kernel\Excel::TYPE_DOUBLE,
        \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
    ])
    ->nextCellCallback(function ($row, $cell, $data) {
        echo 'cell:' . $cell . ', row:' . $row .', data type:' . gettype($data) . PHP_EOL;
    });
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_cell_callback_with_data_type.xlsx');
?>
--EXPECT--
cell:0, row:0, data type:string
cell:1, row:0, data type:string
cell:2, row:0, data type:string
cell:3, row:0, data type:string
cell:4, row:0, data type:string
cell:4, row:0, data type:string
cell:0, row:1, data type:string
cell:1, row:1, data type:string
cell:2, row:1, data type:integer
cell:3, row:1, data type:double
cell:4, row:1, data type:integer
cell:4, row:1, data type:string
