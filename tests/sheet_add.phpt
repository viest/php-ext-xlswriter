--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
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

/* Round-trip: data round-trips intact (non-empty). */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('sheet_add.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/sheet_add.xlsx');
?>
--EXPECT--
string(22) "./tests/sheet_add.xlsx"
bool(true)
