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
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1'],
    ])
    ->output();

$excel->openFile('tutorial.xlsx')->nextCellCallback(function ($row, $cell, $data) {
    echo 'cell:' . $cell . ', row:' . $row . ', value:' . $data . PHP_EOL;
});
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
cell:1, row:1, value:Item
cell:2, row:1, value:Cost
cell:2, row:1, value:XLSX_ROW_END
cell:1, row:2, value:Item_1
cell:2, row:2, value:Cost_1
cell:2, row:2, value:XLSX_ROW_END