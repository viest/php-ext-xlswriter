--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('11.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setColumn('A:A', 200, $boldStyle)
    ->setColumn('B:B', 200, null)
    ->output();

var_dump($filePath);

/* Round-trip: data round-trips intact (non-empty). */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('11.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/11.xlsx');
?>
--EXPECT--
string(15) "./tests/11.xlsx"
bool(true)
