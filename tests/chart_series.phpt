--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('chart_series.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_COLUMN);

$chartResource = $chart->series('Sheet1!$A$1:$A$5')
    ->series('Sheet1!$B$1:$B$5')
    ->series('Sheet1!$C$1:$C$5')
    ->toResource();

var_dump($chartResource);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/chart_series.xlsx');
?>
--EXPECT--
resource(5) of type (xlsx)
