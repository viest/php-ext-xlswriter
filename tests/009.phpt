--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel  = new \Vtiful\Kernel\Excel($config);
$handle = $excel->fileName('tutorial01.xlsx')->getHandle();

$format         = new \Vtiful\Kernel\Format($handle);
$underlineStyle = $format->underline(\Vtiful\Kernel\Format::UNDERLINE_SINGLE)->toResource();

var_dump($underlineStyle);
?>
--EXPECT--
resource(5) of type (xlsx)
