--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('insert_url_format.xlsx');
$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$alignStyle = $format->align(
    \Vtiful\Kernel\Format::FORMAT_ALIGN_CENTER,
    \Vtiful\Kernel\Format::FORMAT_ALIGN_VERTICAL_CENTER
)->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->insertUrl(3, 0, 'https://github.com', NULL, NULL, $alignStyle)
    ->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('insert_url_format.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_url_format.xlsx');
?>
--EXPECT--
string(30) "./tests/insert_url_format.xlsx"
bool(true)
