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

$fileObject = $fileObject->fileName('format_font_size.xlsx');
$fileHandle = $fileObject->getHandle();

$format = new \Vtiful\Kernel\Format($fileHandle);
$style  = $format->fontSize(30)->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setRow('A1', 50, $style)
    ->setRow('A2:A3', 50, $style)
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_font_size.xlsx');
?>
--EXPECT--
string(29) "./tests/format_font_size.xlsx"
