--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php 
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$handle = $excel->fileName('tutorial01.xlsx')
    ->getHandle();
$underlineStyle = \Vtiful\Kernel\Format::underline($handle, \Vtiful\Kernel\Format::UNDERLINE_SINGLE);
var_dump($underlineStyle);
?>
--EXPECT--
resource(5) of type (xlsx)
