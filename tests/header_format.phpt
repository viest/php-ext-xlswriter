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

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('header_format.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['alignment']['horizontal'], $fmt['alignment']['vertical']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/header_format.xlsx');
?>
--EXPECT--
string(26) "./tests/header_format.xlsx"
string(6) "center"
string(6) "center"
