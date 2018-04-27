--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php 
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
var_dump($excel);
?>
--EXPECT--
object(Vtiful\Kernel\Excel)#1 (2) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  NULL
}

