--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests'
];

$fileObject  = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('format_unlocked.xlsx');
$fileHandle = $fileObject->getHandle();

$format       = new \Vtiful\Kernel\Format($fileHandle);
$unlockedStyle = $format->unlocked()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['wjx',   21]
    ])
    ->setRow('A2', 50, $unlockedStyle)
    ->protection()
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_unlocked.xlsx');
?>
--EXPECT--
string(28) "./tests/format_unlocked.xlsx"
