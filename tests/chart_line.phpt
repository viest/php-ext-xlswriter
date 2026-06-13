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

$fileObject = $fileObject->fileName('chart_line.xlsx');
$fileHandle = $fileObject->getHandle();

$chart         = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_LINE);
$chartResource = $chart->series('Sheet1!$A$2:$A$7')->toResource();

$filePath = $fileObject->header(['number'])
    ->data([
        [10],
        [40],
        [50],
        [20],
        [10],
        [50],
    ])
    ->insertChart(0, 3, $chartResource)
    ->output();

var_dump($filePath);

/* Round-trip: chart's underlying data sheet reads back the 6 numbers. */
$v_ = new \Vtiful\Kernel\Excel($config);
$d_ = $v_->openFile('chart_line.xlsx')->openSheet()->getSheetData();
$nums = [];
foreach ($d_ as $row) { foreach ($row as $cell) { $nums[] = (int)$cell; } }
$nums = array_slice($nums, 1);
sort($nums);
var_dump($nums);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/chart_line.xlsx');
?>
--EXPECT--
string(23) "./tests/chart_line.xlsx"
array(6) {
  [0]=>
  int(10)
  [1]=>
  int(10)
  [2]=>
  int(20)
  [3]=>
  int(40)
  [4]=>
  int(50)
  [5]=>
  int(50)
}
