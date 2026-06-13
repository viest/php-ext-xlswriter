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

$fileObject = $fileObject->fileName('format_font_color.xlsx');
$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setRow('A1', 50, $colorStyle)
    ->output();

var_dump($filePath);

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_font_color.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['font']['color']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_font_color.xlsx');
?>
--EXPECT--
string(30) "./tests/format_font_color.xlsx"
string(8) "FFFF6600"
