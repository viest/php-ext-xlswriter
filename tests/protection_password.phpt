--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('protection_password.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection('password')
    ->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('protection_password.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/protection_password.xlsx');
?>
--EXPECT--
string(32) "./tests/protection_password.xlsx"
array(3) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "name"
    [1]=>
    string(3) "age"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    int(21)
  }
  [2]=>
  array(2) {
    [0]=>
    string(3) "wjx"
    [1]=>
    int(21)
  }
}
