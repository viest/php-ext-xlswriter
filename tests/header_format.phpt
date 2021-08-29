--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('header_format.xlsx');
$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$alignStyle = $format->align(
    \Vtiful\Kernel\Format::FORMAT_ALIGN_CENTER,
    \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_CENTER
)->toResource();

$setHeader = $fileObject
    ->header(['Item', 'Cost'], $alignStyle)
    ->output();

var_dump($setHeader);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/header_format.xlsx');
?>
--EXPECT--
string(26) "./tests/header_format.xlsx"