--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
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

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('outline_level_row.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_level_row.xlsx');
?>
--EXPECT--
string(30) "./tests/outline_level_row.xlsx"
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
