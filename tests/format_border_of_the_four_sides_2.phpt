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

$fileObject = $fileObject->fileName('format_border_of_the_four_side_2.xlsx');
$fileHandle = $fileObject->getHandle();

$data = [
    ['viest1', 21, 100, "A"],
    ['viest2', 20, 80, "B"],
    ['viest3', 22, 70, "C"],
];

$format = new \Vtiful\Kernel\Format($fileHandle);

$borderStyle = $format
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->borderOfTheFourSides(
            NULL,                                 // top
            \Vtiful\Kernel\Format::BORDER_THICK,  // right
            \Vtiful\Kernel\Format::BORDER_DOUBLE, // bottom
            NULL                                  // left
        )
    ->toResource();

$filePath = $fileObject->header(['name', 'age', 'score', 'level'])
    ->data($data)
    ->setRow('A1', 20, $borderStyle)
    ->output();

var_dump($filePath);

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_border_of_the_four_side_2.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['border']['right']['style'], $fmt['border']['bottom']['style']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_border_of_the_four_side_2.xlsx');
?>
--EXPECT--
string(45) "./tests/format_border_of_the_four_side_2.xlsx"
string(5) "thick"
string(6) "double"
