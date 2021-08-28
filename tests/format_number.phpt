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

$fileObject = $fileObject->fileName('format_number.xlsx');
$fileHandle = $fileObject->getHandle();

$format      = new \Vtiful\Kernel\Format($fileHandle);
$numberStyle = $format->number('#,##0')->toResource();

$filePath = $fileObject->header(['name', 'balance'])
    ->data([
        ['viest', 10000],
        ['wjx',   100000]
    ])
    ->setColumn('B:B', 50, $numberStyle)
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_number.xlsx');
?>
--EXPECT--
string(26) "./tests/format_number.xlsx"
