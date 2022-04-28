--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$data = [
    [23],
    [21],
    [21],
    [21],
];

foreach($data as &$line) {
    $line[0]++;
}

$excel = new \Vtiful\Kernel\Excel([
    'path' => './tests',
]);

$fileObject = $excel->constMemory('data_reference.xlsx', NULL, false);
$fileHandle = $fileObject->getHandle();

$path = $fileObject->header(['age'])
    ->data($data)
    ->output();

$excel->openFile('data_reference.xlsx')
    ->openSheet();

var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/data_reference.xlsx');
?>
--EXPECT--
array(1) {
  [0]=>
  string(3) "age"
}
array(1) {
  [0]=>
  int(24)
}
array(1) {
  [0]=>
  int(22)
}
array(1) {
  [0]=>
  int(22)
}
