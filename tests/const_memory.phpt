--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$excel = new \Vtiful\Kernel\Excel([
    'path' => './tests',
]);

$fileObject = $excel->constMemory('const_memory.xlsx');
$fileHandle = $fileObject->getHandle();

$path = $fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->output();

var_dump($path);

/* Round-trip: const-memory writer produces the same bytes as the full
   in-memory writer for header + one data row. */
$v_ = new \Vtiful\Kernel\Excel(['path' => './tests']);
$d_ = $v_->openFile('const_memory.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/const_memory.xlsx');
?>
--EXPECT--
string(25) "./tests/const_memory.xlsx"
array(2) {
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
}