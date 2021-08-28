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
$filePath = $excel->fileName('open_xlsx_next_row_skip_empty.xlsx')
    ->header(['', 'Cost'])
    ->data([
        [],
        ['viest', ''],
    ])
    ->output();

echo 'skip cells' . PHP_EOL;

$data = $excel->openFile('open_xlsx_next_row_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS);

while (is_array($data = $excel->nextRow())) {
    var_dump($data);
}

echo 'skip row' . PHP_EOL;

$data = $excel->openFile('open_xlsx_next_row_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW);

while (is_array($data = $excel->nextRow())) {
    var_dump($data);
}

echo 'skip cells & row' . PHP_EOL;

$data = $excel->openFile('open_xlsx_next_row_skip_empty.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS | \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW);

while (is_array($data = $excel->nextRow())) {
    var_dump($data);
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row_skip_empty.xlsx');
?>
--EXPECT--
skip cells
array(1) {
  [1]=>
  string(4) "Cost"
}
array(0) {
}
array(1) {
  [0]=>
  string(5) "viest"
}
skip row
array(2) {
  [0]=>
  string(0) ""
  [1]=>
  string(4) "Cost"
}
array(2) {
  [0]=>
  string(5) "viest"
  [1]=>
  string(0) ""
}
skip cells & row
array(1) {
  [1]=>
  string(4) "Cost"
}
array(1) {
  [0]=>
  string(5) "viest"
}