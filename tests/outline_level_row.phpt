--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/*
 * https://libxlsxwriter.github.io/working_with_outlines.html
 */
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('outline_level_row.xlsx');

$format    = new \Vtiful\Kernel\Format($excel->getHandle());
$boldStyle = $format->bold()->toResource();

$filePath = $excel
    ->header(['Region', 'Sales'])
    ->data([
        ['North', 1000], // row 2
        ['North', 1200],
        ['North', 900],
        ['North', 1200],
        ['North Total', 4300], // row 6
        ['South', 400], // row 7
        ['South', 600],
        ['South', 500],
        ['South', 600],
        ['South Total', 2100], // row 11
        ['Grand Total', 6400], // row 12
        ['hidden row', 0],
    ])
    ->setRow('A1', 15, $boldStyle)
    ->setRow('A2:A5', 15, null, 2, false, true)
    ->setRow('A6', 15, $boldStyle, 1, true)
    ->setRow('A7:A10', 15, null, 2)
    ->setRow('A11', 15, $boldStyle, 1)
    ->setRow('A12', 15, $boldStyle, 0)
    ->setRow('A13', 15, $boldStyle, null, null, true)
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_level_row.xlsx');
?>
--EXPECT--
string(30) "./tests/outline_level_row.xlsx"
