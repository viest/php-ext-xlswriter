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

$fileObject = $fileObject->fileName('chart_resource.xlsx');
$fileHandle = $fileObject->getHandle();

$chart         = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_LINE);
$chartResource = $chart->toResource();

var_dump($chartResource);
?>
--CLEAN--
<?php
//
?>
--EXPECT--
resource(5) of type (xlsx)
