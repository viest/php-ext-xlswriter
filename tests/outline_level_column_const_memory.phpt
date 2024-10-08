--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/*
 * https://libxlsxwriter.github.io/working_with_outlines.html
 * Same result as in outline_level_column, but with const memory.
 */
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->constMemory('outline_level_column_const_memory.xlsx');

$format       = new \Vtiful\Kernel\Format($excel->getHandle());
$defaultStyle = $format->toResource();

$format    = new \Vtiful\Kernel\Format($excel->getHandle());
$boldStyle = $format->bold()->toResource();

$filePath = $excel
    ->defaultFormat($boldStyle)
    ->header(['Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total', 'hidden column I'])
    ->defaultFormat($defaultStyle)
    ->data([
        ['North', 50, 20, 15, 25, 65, 80, 255],
        ['South', 10, 20, 30, 50, 50, 50, 210],
        ['East', 45, 75, 50, 15, 75, 100, 360],
        ['West', 15, 15, 55, 35, 20, 50, 190],
    ])
    ->insertText(5, 7, 1015, '', $boldStyle)
    ->setColumn('A:A', 10, $boldStyle)
    ->setColumn('B:G', 10, null, 1)
    ->setColumn('H:H', 10, null, 0, true)
    ->setColumn('I:I', 10, null, null, false, true)
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_level_column_const_memory.xlsx');
?>
--EXPECT--
string(46) "./tests/outline_level_column_const_memory.xlsx"
