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
$filePath = $excel->fileName('open_xlsx_next_cell_callback.xlsx')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1'],
    ])
    ->output();

$excel->openFile('open_xlsx_next_cell_callback.xlsx')->nextCellCallback(function ($row, $cell, $data) {
    echo 'cell:' . $cell . ', row:' . $row . ', value:' . $data . PHP_EOL;
});
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_cell_callback.xlsx');
?>
--EXPECT--
cell:0, row:0, value:Item
cell:1, row:0, value:Cost
cell:1, row:0, value:XLSX_ROW_END
cell:0, row:1, value:Item_1
cell:1, row:1, value:Cost_1
cell:1, row:1, value:XLSX_ROW_END