--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$fileFd = $excel->fileName('004.xlsx');
$setHeader = $fileFd->header(['Item', 'Cost']);
var_dump($setHeader);
?>
--CLEAN--
<?php
@unlink(__DIR__ . "/004.xlsx");
?>
--EXPECTF--
object(Vtiful\Kernel\Excel)#%d (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(16) "./tests/004.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}