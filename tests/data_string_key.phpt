--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$excel = new \Vtiful\Kernel\Excel([
    'path' => './tests',
]);

$fileObject = $excel->constMemory('data_string_key.xlsx', NULL, false);
$fileHandle = $fileObject->getHandle();

$path = $fileObject->header(['name', 'age'])
    ->data([
        ['name'=>'viest', 'age' => 21, 1 => 23],
        ['name'=>'viest', 'age' => 21],
        ['name'=>'viest', 'age' => 21],
        ['viest', 21],
    ])
    ->output();

$excel->openFile('data_string_key.xlsx')
    ->openSheet();

var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
var_dump($excel->nextRow());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/data_string_key.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  string(4) "name"
  [1]=>
  string(3) "age"
}
array(2) {
  [0]=>
  string(5) "viest"
  [1]=>
  int(23)
}
array(2) {
  [0]=>
  string(5) "viest"
  [1]=>
  int(21)
}
array(2) {
  [0]=>
  string(5) "viest"
  [1]=>
  int(21)
}
