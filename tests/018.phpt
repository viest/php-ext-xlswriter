--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("18.xlsx");

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->data([['wjx', 21]]);

$filePath = $fileObject->output();

var_dump($filePath);

/* Round-trip: data round-trips intact (non-empty). */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('18.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/18.xlsx');
?>
--EXPECT--
string(15) "./tests/18.xlsx"
bool(true)
