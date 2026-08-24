--TEST--
mergeCells() with a single-cell range writes the value in edit mode too
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

// Source file already has a merged range A1:C1.
(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_merge_single_cell_src.xlsx', 'S')
    ->mergeCells('A1:C1', 'title')
    ->output();

// 'A3:A3' covers one cell, which Excel cannot merge: editing writes the value
// without a merge record rather than raising an error.
(new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_merge_single_cell_src.xlsx')
    ->openSheet('S')
    ->mergeCells('A3:A3', 'solo')
    ->output('edit_merge_single_cell_out.xlsx');

// Only the pre-existing merge survives.
$merges = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_merge_single_cell_out.xlsx')
    ->openSheet('S')
    ->getMergedCells();

foreach ($merges as $m) {
    echo "{$m['first_row']},{$m['first_col']},{$m['last_row']},{$m['last_col']}", PHP_EOL;
}

$data = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_merge_single_cell_out.xlsx')
    ->openSheet('S')
    ->getSheetData();

var_dump($data[2][0]);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_merge_single_cell_src.xlsx');
@unlink(__DIR__ . '/edit_merge_single_cell_out.xlsx');
?>
--EXPECT--
1,1,1,3
string(4) "solo"
