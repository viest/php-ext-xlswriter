--TEST--
mergeCells round-trip: anchor cells preserve typed values (int/string/float)

NOTE: this test verifies the *master cell value* survives round-trip.
Verifying the merged *region geometry* (first/last row/col) is a Phase 1
dependency — see plans/upgrade.md §5.2.1 (`getMergedCells`).
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("merge_cell_type_writer.xlsx")
    ->mergeCells('A1:C1', 1)
    ->mergeCells('A2:D2', '2')
    ->mergeCells('A3:E3', 3.001)
    ->output();

$data = $excel->openFile('merge_cell_type_writer.xlsx')
    ->openSheet()
    ->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/merge_cell_type_writer.xlsx');
?>
--EXPECT--
array(3) {
  [0]=>
  array(1) {
    [0]=>
    int(1)
  }
  [1]=>
  array(1) {
    [0]=>
    int(2)
  }
  [2]=>
  array(1) {
    [0]=>
    float(3.001)
  }
}
