--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests',
];

$lastFilePath = NULL;

for ($index = 0; $index < 100; $index++) {
    $fileObject = new \Vtiful\Kernel\Excel($config);

    $fileObject = $fileObject->fileName('multiple_file' . $index . '.xlsx');
    $fileHandle = $fileObject->getHandle();

    $format     = new \Vtiful\Kernel\Format($fileHandle);
    $alignStyle = $format->align(
        \Vtiful\Kernel\Format::FORMAT_ALIGN_CENTER,
        \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_CENTER
    )->toResource();

    $lastFilePath = $fileObject->header(['name', 'age'])
        ->data([
            ['viest', 21],
            ['wjx', 21],
        ])
        ->setRow('A1', 50, $alignStyle)
        ->setRow('A2:A3', 50, $alignStyle)
        ->output();
}

var_dump($lastFilePath);
?>
--CLEAN--
<?php
for ($index = 0; $index < 100; $index++) {
    @unlink(__DIR__ . '/multiple_file' . $index . '.xlsx');
}
?>
--EXPECT--
string(28) "./tests/multiple_file99.xlsx"
