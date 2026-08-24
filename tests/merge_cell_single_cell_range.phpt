--TEST--
mergeCells() with a single-cell range writes the value instead of failing
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

// 'C1:C1' covers one cell, which Excel cannot merge: the value is written
// without a merge record rather than raising an error.
$excel->fileName('merge_cell_single_cell_range.xlsx', 'Sheet1')
    ->header(['Item', 'Cost'])
    ->mergeCells('C1:C1', 'Days')
    ->mergeCells('D1:E1', 'Note')
    ->data([
        ['Rent', 1000, 1],
        ['Gas', 100, 5],
    ])
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('merge_cell_single_cell_range.xlsx')
    ->openSheet();

var_dump($reader->getMergedCells());

$data = (new \Vtiful\Kernel\Excel($config))
    ->openFile('merge_cell_single_cell_range.xlsx')
    ->openSheet()
    ->getSheetData();

var_dump($data[0]);

// A genuinely invalid range must still raise: degrading the single-cell case
// may not swallow the other error codes.
try {
    (new \Vtiful\Kernel\Excel($config))
        ->fileName('merge_cell_single_cell_range_err.xlsx', 'Sheet1')
        ->mergeCells('A1:B1048577', 'boom');
} catch (\Vtiful\Kernel\Exception $e) {
    echo $e->getMessage(), PHP_EOL;
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/merge_cell_single_cell_range.xlsx');
@unlink(__DIR__ . '/merge_cell_single_cell_range_err.xlsx');
?>
--EXPECT--
array(1) {
  [0]=>
  array(4) {
    ["first_row"]=>
    int(1)
    ["first_col"]=>
    int(4)
    ["last_row"]=>
    int(1)
    ["last_col"]=>
    int(5)
  }
}
array(4) {
  [0]=>
  string(4) "Item"
  [1]=>
  string(4) "Cost"
  [2]=>
  string(4) "Days"
  [3]=>
  string(4) "Note"
}
Worksheet row or column index out of range
