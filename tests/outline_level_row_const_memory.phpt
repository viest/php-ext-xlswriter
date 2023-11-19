--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/*
 * https://libxlsxwriter.github.io/working_with_outlines.html
 * Same result as in outline_level_row, but with const memory.
 */
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->constMemory('outline_level_row_const_memory.xlsx');

$format       = new \Vtiful\Kernel\Format($excel->getHandle());
$defaultStyle = $format->toResource();

$format    = new \Vtiful\Kernel\Format($excel->getHandle());
$boldStyle = $format->bold()->toResource();

$filePath = $excel
    ->defaultFormat($boldStyle)
    ->header(['Region', 'Sales'])

    ->defaultFormat($defaultStyle)
    ->defaultRowOptions(2, false, true)
    ->data([
        ['North', 1000],
        ['North', 1200],
        ['North', 900],
        ['North', 1200],
    ])

    ->defaultFormat($boldStyle)
    ->defaultRowOptions(1, true)
    ->data([
        ['North Total', 4300],
    ])

    ->defaultFormat($defaultStyle)
    ->defaultRowOptions(2)
    ->data([
        ['South', 400],
        ['South', 600],
        ['South', 500],
        ['South', 600],
    ])

    ->defaultFormat($boldStyle)
    ->defaultRowOptions(1)
    ->data([
        ['South Total', 2100],
    ])

    ->defaultRowOptions(0)
    ->data([
        ['Grand Total', 6400],
    ])

    ->defaultRowOptions(null, null, true)
    ->data([
        ['hidden row', 0],
    ])

    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_level_row_const_memory.xlsx');
?>
--EXPECT--
string(43) "./tests/outline_level_row_const_memory.xlsx"
