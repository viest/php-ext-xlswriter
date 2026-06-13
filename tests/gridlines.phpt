--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("gridlines.xlsx");

$fileObject->header(['name', 'age'])
    ->gridline(\Vtiful\Kernel\Excel::GRIDLINES_HIDE_ALL)
    ->data([
        ['viest', 21],
        ['viest', 22],
        ['viest', 23],
    ]);

$filePath = $fileObject->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('gridlines.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/gridlines.xlsx');
?>
--EXPECT--
string(22) "./tests/gridlines.xlsx"
array(4) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "name"
    [1]=>
    string(3) "age"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    int(21)
  }
  [2]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    int(22)
  }
  [3]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    int(23)
  }
}
