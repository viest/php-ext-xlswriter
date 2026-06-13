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

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_rotation.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['alignment']['rotation']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_rotation.xlsx');
?>
--EXPECT--
string(28) "./tests/format_rotation.xlsx"
int(30)
