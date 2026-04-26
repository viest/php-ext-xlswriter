--TEST--
SKIP_MERGED_FOLLOW: non-master cells inside merges read back as null
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_skip_merged_follow.xlsx', 'S1')
    ->insertText(0, 0, 'A')
    ->insertText(0, 1, 'B')
    ->insertText(0, 2, 'C')
    ->mergeCells('A2:C2', 'M')
    ->insertText(2, 0, 'below')
    ->output();

echo "default:\n";
$r1 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_skip_merged_follow.xlsx')
    ->openSheet();
while ($row = $r1->nextRow()) var_dump($row);

echo "skip merged follow:\n";
$r2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_skip_merged_follow.xlsx')
    ->openSheet(null, \Vtiful\Kernel\Excel::SKIP_MERGED_FOLLOW);
while ($row = $r2->nextRow()) var_dump($row);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_skip_merged_follow.xlsx');
?>
--EXPECT--
default:
array(3) {
  [0]=>
  string(1) "A"
  [1]=>
  string(1) "B"
  [2]=>
  string(1) "C"
}
array(3) {
  [0]=>
  string(1) "M"
  [1]=>
  string(0) ""
  [2]=>
  string(0) ""
}
array(3) {
  [0]=>
  string(5) "below"
  [1]=>
  string(0) ""
  [2]=>
  string(0) ""
}
skip merged follow:
array(3) {
  [0]=>
  string(1) "A"
  [1]=>
  string(1) "B"
  [2]=>
  string(1) "C"
}
array(3) {
  [0]=>
  string(1) "M"
  [1]=>
  NULL
  [2]=>
  NULL
}
array(3) {
  [0]=>
  string(5) "below"
  [1]=>
  string(0) ""
  [2]=>
  string(0) ""
}
