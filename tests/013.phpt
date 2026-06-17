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

$textFile = $excel->fileName("13.xlsx")
    ->header(['name', 'age']);

for ($index = 0; $index < 10; $index++) {
    $textFile->insertText($index+1, 0, 'vikin');
    $textFile->insertText($index+1, 1, 1000, '#,##0');
}

$filePath = $textFile->output();

var_dump($filePath);

/* Round-trip: 10 rows of insertText (with #,##0 format) read back as ints. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('13.xlsx')->openSheet()->getSheetData();
var_dump(count($d_));
var_dump($d_[0]);
var_dump($d_[1]);
var_dump($d_[10]);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/13.xlsx');
?>
--EXPECT--
string(15) "./tests/13.xlsx"
int(11)
array(2) {
  [0]=>
  string(4) "name"
  [1]=>
  string(3) "age"
}
array(2) {
  [0]=>
  string(5) "vikin"
  [1]=>
  int(1000)
}
array(2) {
  [0]=>
  string(5) "vikin"
  [1]=>
  int(1000)
}
