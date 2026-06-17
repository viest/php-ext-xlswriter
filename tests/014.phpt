--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("14.xlsx")
    ->header(['name', 'money']);

for($index = 1; $index <= 10; $index++) {
    $freeFile->insertText($index, 0, 'vikin');
    $freeFile->insertText($index, 1, 10);
}

$freeFile->insertText(12, 0, "Total");
$freeFile->insertFormula(12, 1, '=SUM(B2:B11)');
$freeFile->insertText(13, 0, "Total (default format)");
$freeFile->insertFormula(13, 1, '=SUM(B2:B11)', null);

$filePath = $freeFile->output();

var_dump($filePath);

/* Round-trip: 10 data rows, a synthesised blank, then two SUM rows.
   Formula cells round-trip as the cached string value libxlsxwriter writes
   ("0"), since the writer doesn't evaluate the formula. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('14.xlsx')->openSheet()->getSheetData();
var_dump(count($d_));
var_dump($d_[0]);
var_dump($d_[1]);
var_dump($d_[10]);
var_dump($d_[11]);
var_dump($d_[12]);
var_dump($d_[13]);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/14.xlsx');
?>
--EXPECT--
string(15) "./tests/14.xlsx"
int(14)
array(2) {
  [0]=>
  string(4) "name"
  [1]=>
  string(5) "money"
}
array(2) {
  [0]=>
  string(5) "vikin"
  [1]=>
  int(10)
}
array(2) {
  [0]=>
  string(5) "vikin"
  [1]=>
  int(10)
}
array(0) {
}
array(2) {
  [0]=>
  string(5) "Total"
  [1]=>
  string(1) "0"
}
array(2) {
  [0]=>
  string(22) "Total (default format)"
  [1]=>
  string(1) "0"
}
