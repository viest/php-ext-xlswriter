--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject->fileName('activate_sheet.xlsx')
    ->header(['name', 'age'])
    ->data([
    ['viest', 21],
    ['viest', 22],
    ['viest', 23],
    ]);

$fileObject->addSheet('twoSheet')
    ->header(['name', 'age'])
    ->data([['vikin', 22]]);

var_dump($fileObject->activateSheet('twoSheet'));

$fileObject->output();

/* Round-trip: both sheets present in the workbook. */
$v_      = new \Vtiful\Kernel\Excel($config);
$sheets_ = $v_->openFile('activate_sheet.xlsx')->sheetList();
sort($sheets_);
var_dump($sheets_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/activate_sheet.xlsx');
?>
--EXPECT--
bool(true)
array(2) {
  [0]=>
  string(6) "Sheet1"
  [1]=>
  string(8) "twoSheet"
}
