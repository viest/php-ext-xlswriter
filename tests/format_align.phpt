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

$fileObject = $fileObject->fileName('format_align.xlsx');
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
    ->setRow('A1', 50, $alignStyle)
    ->setRow('A2:A3', 50, $alignStyle)
    ->output();

var_dump($filePath);

/* Round-trip: the styled cell records the exact horizontal+vertical
 * alignment we wrote. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_align.xlsx')->openSheet();
$verify->nextRowWithFormula();          /* header */
$row    = $verify->nextRowWithFormula();
$sid    = 0;
foreach ($row as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break; } }
$fmt    = $verify->getStyleFormat($sid);
var_dump($fmt['alignment']['horizontal'], $fmt['alignment']['vertical']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_align.xlsx');
?>
--EXPECT--
string(25) "./tests/format_align.xlsx"
string(6) "center"
string(6) "center"
