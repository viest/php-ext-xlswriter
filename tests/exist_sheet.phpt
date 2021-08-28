--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject->fileName('exist_sheet.xlsx')
    ->addSheet('twoSheet');

var_dump($fileObject->existSheet('twoSheet'));
var_dump($fileObject->existSheet('notFoundSheet'));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/exist_sheet.xlsx');
?>
--EXPECT--
bool(true)
bool(false)
