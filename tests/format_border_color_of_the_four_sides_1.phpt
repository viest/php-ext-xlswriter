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

$fileObject = $fileObject->fileName('format_border_color_of_the_four_side_1.xlsx');
$fileHandle = $fileObject->getHandle();

$data = [
    ['viest1', 21, 100, "A"],
    ['viest2', 20, 80, "B"],
    ['viest3', 22, 70, "C"],
];

$format = new \Vtiful\Kernel\Format($fileHandle);

$borderStyle = $format
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->borderColorOfTheFourSides(
        \Vtiful\Kernel\Format::COLOR_ORANGE, // top
        \Vtiful\Kernel\Format::COLOR_GREEN,  // right
        \Vtiful\Kernel\Format::COLOR_RED,    // bottom
        \Vtiful\Kernel\Format::COLOR_YELLOW  // left
    )
    ->toResource();

$filePath = $fileObject->header(['name', 'age', 'score', 'level'])
    ->data($data)
    ->setRow('A1', 20, $borderStyle)
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_border_color_of_the_four_side_1.xlsx');
?>
--EXPECT--
string(51) "./tests/format_border_color_of_the_four_side_1.xlsx"
