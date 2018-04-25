--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("excel_writer")) print "skip"; ?>
--FILE--
<?php 
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$handle = $excel->fileName('tutorial01.xlsx')
    ->getHandle();
$boldStyle = \Vtiful\Kernel\Format::bold($handle);
var_dump($boldStyle);
?>
--EXPECT--
resource(5) of type (excel)
