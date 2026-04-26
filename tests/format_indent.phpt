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

$fileObject = $fileObject->fileName('format_indent.xlsx');
$fileHandle = $fileObject->getHandle();

$format = new \Vtiful\Kernel\Format($fileHandle);
$indentStyle = $format->indent(1)->toResource();

$fileObject->insertText(0, 0, 'viest', null, $indentStyle);
$fileObject->insertText(0, 1, 'wjx');

$filePath = $fileObject->output();

var_dump($filePath);

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_indent.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['alignment']['horizontal'], $fmt['alignment']['indent']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_indent.xlsx');
?>
--EXPECT--
string(26) "./tests/format_indent.xlsx"
string(4) "left"
int(1)
