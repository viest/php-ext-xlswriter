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

$fileObject = $excel->fileName("sheet_checkout.xlsx");

$fileObject->header(['name', 'age'])
    ->data([
    ['viest', 21],
    ['viest', 22],
    ['viest', 23],
    ]);

$fileObject->addSheet('twoSheet')
    ->header(['name', 'age'])
    ->data([['vikin', 22]]);

$fileObject->checkoutSheet('Sheet1')
    ->data([['sheet1']]);

$filePath = $fileObject->output();

var_dump($filePath);

/* Round-trip: checkoutSheet('Sheet1') resumes appending after the second
   ->data() block landed on twoSheet, so 'sheet1' ends up as the 4th row
   on Sheet1 (after the 3 viest rows). The second column there is empty. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('sheet_checkout.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/sheet_checkout.xlsx');
?>
--EXPECT--
string(27) "./tests/sheet_checkout.xlsx"
array(5) {
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
    string(5) "viest"
    [1]=>
    int(22)
  }
  [3]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    int(23)
  }
  [4]=>
  array(2) {
    [0]=>
    string(6) "sheet1"
    [1]=>
    string(0) ""
  }
}
