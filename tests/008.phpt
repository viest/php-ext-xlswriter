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
$italicStyle = \Vtiful\Kernel\Format::italic($handle);
var_dump($italicStyle);
?>
--EXPECT--
resource(5) of type (vtiful)
