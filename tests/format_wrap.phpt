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

$fileObject = $fileObject->fileName('format_wrap.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$wrapStyle = $format->wrap()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ["vvvvvvvvvvvvvvvvvvvvvvvvvv\nvvvvvvvvvvvvvvvvvvvvvvvvvvv", 21],
        ['wjx',   21]
    ])
    ->setRow('A2', 50, $wrapStyle)
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_wrap.xlsx');
?>
--EXPECT--
string(24) "./tests/format_wrap.xlsx"
