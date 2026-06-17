--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = [
    'path' => './tests',
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('freeze_panes.xlsx');
$fileHandle = $fileObject->getHandle();

$filePath = $fileObject->freezePanes(1, 0)
    ->header(['name', 'age']);

for ($index = 0; $index < 100; $index++) {
    $fileObject->insertText($index + 1, 0, 'wjx');
    $fileObject->insertText($index + 1, 1, 21);
}

var_dump($fileObject->output());

/* Round-trip: 1 header + 100 data rows, sampled at both ends. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('freeze_panes.xlsx')->openSheet()->getSheetData();
var_dump(count($d_));
var_dump($d_[0]);
var_dump($d_[1]);
var_dump($d_[100]);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/freeze_panes.xlsx');
?>
--EXPECT--
string(25) "./tests/freeze_panes.xlsx"
int(101)
array(2) {
  [0]=>
  string(4) "name"
  [1]=>
  string(3) "age"
}
array(2) {
  [0]=>
  string(3) "wjx"
  [1]=>
  int(21)
}
array(2) {
  [0]=>
  string(3) "wjx"
  [1]=>
  int(21)
}
