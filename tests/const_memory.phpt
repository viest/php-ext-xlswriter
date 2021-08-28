--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$excel = new \Vtiful\Kernel\Excel([
    'path' => './tests',
]);

$fileObject = $excel->constMemory('const_memory.xlsx');
$fileHandle = $fileObject->getHandle();

$path = $fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->output();

var_dump($path);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/const_memory.xlsx');
?>
--EXPECT--
string(25) "./tests/const_memory.xlsx"