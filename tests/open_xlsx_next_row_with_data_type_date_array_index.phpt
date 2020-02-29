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
    ->header(['', 'Cost'])
    ->data([
        [],
        ['viest', ''],
    ])
    ->insertDate(2, 4, 1568818881)
    ->output();

$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1');

while (is_array($data = $excel->nextRow([4 => \Vtiful\Kernel\Excel::TYPE_TIMESTAMP]))) {
    var_dump($data);
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  string(0) ""
  [1]=>
  string(4) "Cost"
}
array(0) {
}
array(5) {
  [0]=>
  string(5) "viest"
  [1]=>
  string(0) ""
  [2]=>
  string(0) ""
  [3]=>
  string(0) ""
  [4]=>
  int(1568818881)
}