--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("sheet_add.xlsx");

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]]);

$fileObject->addSheet('twoSheet')
    ->header(['name', 'age'])
    ->data([['vikin', 22]]);

$filePath = $fileObject->output();

var_dump($filePath);

/* Round-trip: both sheets carry their own header + row. */
$v_   = new \Vtiful\Kernel\Excel($config);
$v_->openFile('sheet_add.xlsx');
var_dump($v_->openSheet('Sheet1')->getSheetData());
var_dump($v_->openSheet('twoSheet')->getSheetData());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/sheet_add.xlsx');
?>
--EXPECT--
string(22) "./tests/sheet_add.xlsx"
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
    string(5) "vikin"
    [1]=>
    int(22)
  }
}
