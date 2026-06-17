--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
/*
 * https://libxlsxwriter.github.io/working_with_outlines.html
 */
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('outline_level_column.xlsx');

$format    = new \Vtiful\Kernel\Format($excel->getHandle());
$boldStyle = $format->bold()->toResource();

$filePath = $excel
    ->header(['Month', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Total', 'hidden column I'])
    ->data([
        ['North', 50, 20, 15, 25, 65, 80, 255],
        ['South', 10, 20, 30, 50, 50, 50, 210],
        ['East', 45, 75, 50, 15, 75, 100, 360],
        ['West', 15, 15, 55, 35, 20, 50, 190],
    ])
    ->insertText(5, 7, 1015, '', $boldStyle)
    ->setRow('A1', 15, $boldStyle)
    ->setColumn('A:A', 10, $boldStyle)
    ->setColumn('B:G', 10, null, 1)
    ->setColumn('H:H', 10, null, 0, true)
    ->setColumn('I:I', 10, null, null, false, true)
    ->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('outline_level_column.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_level_column.xlsx');
?>
--EXPECT--
string(33) "./tests/outline_level_column.xlsx"
array(6) {
  [0]=>
  array(9) {
    [0]=>
    string(5) "Month"
    [1]=>
    string(3) "Jan"
    [2]=>
    string(3) "Feb"
    [3]=>
    string(3) "Mar"
    [4]=>
    string(3) "Apr"
    [5]=>
    string(3) "May"
    [6]=>
    string(3) "Jun"
    [7]=>
    string(5) "Total"
    [8]=>
    string(15) "hidden column I"
  }
  [1]=>
  array(9) {
    [0]=>
    string(5) "North"
    [1]=>
    int(50)
    [2]=>
    int(20)
    [3]=>
    int(15)
    [4]=>
    int(25)
    [5]=>
    int(65)
    [6]=>
    int(80)
    [7]=>
    int(255)
    [8]=>
    string(0) ""
  }
  [2]=>
  array(9) {
    [0]=>
    string(5) "South"
    [1]=>
    int(10)
    [2]=>
    int(20)
    [3]=>
    int(30)
    [4]=>
    int(50)
    [5]=>
    int(50)
    [6]=>
    int(50)
    [7]=>
    int(210)
    [8]=>
    string(0) ""
  }
  [3]=>
  array(9) {
    [0]=>
    string(4) "East"
    [1]=>
    int(45)
    [2]=>
    int(75)
    [3]=>
    int(50)
    [4]=>
    int(15)
    [5]=>
    int(75)
    [6]=>
    int(100)
    [7]=>
    int(360)
    [8]=>
    string(0) ""
  }
  [4]=>
  array(9) {
    [0]=>
    string(4) "West"
    [1]=>
    int(15)
    [2]=>
    int(15)
    [3]=>
    int(55)
    [4]=>
    int(35)
    [5]=>
    int(20)
    [6]=>
    int(50)
    [7]=>
    int(190)
    [8]=>
    string(0) ""
  }
  [5]=>
  array(9) {
    [0]=>
    string(0) ""
    [1]=>
    string(0) ""
    [2]=>
    string(0) ""
    [3]=>
    string(0) ""
    [4]=>
    string(0) ""
    [5]=>
    string(0) ""
    [6]=>
    string(0) ""
    [7]=>
    int(1015)
    [8]=>
    string(0) ""
  }
}
