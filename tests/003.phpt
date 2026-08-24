--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$fileFd = $excel->fileName('003.xlsx');
var_dump($fileFd);
?>
--CLEAN--
<?php
@unlink(__DIR__ . "/003.xlsx");
?>
--EXPECTF--
object(Vtiful\Kernel\Excel)#%d (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(16) "./tests/003.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}