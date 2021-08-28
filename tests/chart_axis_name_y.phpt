--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('chart_axis_name_y.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_AREA);

$chartResource = $chart
    ->series('=Sheet1!$B$2:$B$7', '=Sheet1!$A$2:$A$7')
    ->axisNameY('Sample length (mm)')
    ->toResource();

var_dump($chartResource);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/chart_axis_name_y.xlsx');
?>
--EXPECT--
resource(5) of type (xlsx)
