--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel  = new \Vtiful\Kernel\Excel($config);
$handle = $excel->fileName('tutorial01.xlsx')->getHandle();

$format     = new \Vtiful\Kernel\Format($handle);
$boldFormat = $format->bold()->toResource();

var_dump($boldFormat);
?>
--CLEAN--
<?php
@unlink(__DIR__ . "/tutorial01.xlsx");
?>
--EXPECTF--
resource(%d) of type (xlsx)
