--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('12.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21],
        ['abc',   21]
    ])
    ->setRow('A1', 200, $boldStyle)
    ->setRow('A2:A3', 200, $boldStyle)
    ->setRow('A4:A4', 200, null)
    ->output();

var_dump($filePath);

/* Round-trip: header + three data rows read back exactly. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('12.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/12.xlsx');
?>
--EXPECT--
string(15) "./tests/12.xlsx"
array(4) {
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
  [3]=>
  array(2) {
    [0]=>
    string(3) "abc"
    [1]=>
    int(21)
  }
}
