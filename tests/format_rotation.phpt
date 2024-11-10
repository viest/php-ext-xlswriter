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

$fileObject = $fileObject->fileName('format_rotation.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$rotationStyle = $format->rotation(30)->toResource();

$fileObject->insertText(0, 0, 'viest', null, $rotationStyle);
$fileObject->insertText(0, 1, 'wjx');

$filePath = $fileObject->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_rotation.xlsx');
?>
--EXPECT--
string(28) "./tests/format_rotation.xlsx"
