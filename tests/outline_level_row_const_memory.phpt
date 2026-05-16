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

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('outline_level_row_const_memory.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_level_row_const_memory.xlsx');
?>
--EXPECT--
string(43) "./tests/outline_level_row_const_memory.xlsx"
array(13) {
  [0]=>
  array(2) {
    [0]=>
    string(6) "Region"
    [1]=>
    string(5) "Sales"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "North"
    [1]=>
    int(1000)
  }
  [2]=>
  array(2) {
    [0]=>
    string(5) "North"
    [1]=>
    int(1200)
  }
  [3]=>
  array(2) {
    [0]=>
    string(5) "North"
    [1]=>
    int(900)
  }
  [4]=>
  array(2) {
    [0]=>
    string(5) "North"
    [1]=>
    int(1200)
  }
  [5]=>
  array(2) {
    [0]=>
    string(11) "North Total"
    [1]=>
    int(4300)
  }
  [6]=>
  array(2) {
    [0]=>
    string(5) "South"
    [1]=>
    int(400)
  }
  [7]=>
  array(2) {
    [0]=>
    string(5) "South"
    [1]=>
    int(600)
  }
  [8]=>
  array(2) {
    [0]=>
    string(5) "South"
    [1]=>
    int(500)
  }
  [9]=>
  array(2) {
    [0]=>
    string(5) "South"
    [1]=>
    int(600)
  }
  [10]=>
  array(2) {
    [0]=>
    string(11) "South Total"
    [1]=>
    int(2100)
  }
  [11]=>
  array(2) {
    [0]=>
    string(11) "Grand Total"
    [1]=>
    int(6400)
  }
  [12]=>
  array(2) {
    [0]=>
    string(10) "hidden row"
    [1]=>
    string(1) "0"
  }
}
