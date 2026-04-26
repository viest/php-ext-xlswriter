--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
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

/* Round-trip: data round-trips intact (non-empty). */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('sheet_checkout.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/sheet_checkout.xlsx');
?>
--EXPECT--
string(27) "./tests/sheet_checkout.xlsx"
bool(true)
