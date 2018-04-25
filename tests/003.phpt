--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("excel_writer")) print "skip"; ?>
--FILE--
<?php 
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$fileFd = $excel->fileName('tutorial01.xlsx');
var_dump($fileFd);
?>
--EXPECT--
object(Vtiful\Kernel\Excel)#1 (3) {
  ["config"]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName"]=>
  string(23) "./tests/tutorial01.xlsx"
  ["handle"]=>
  resource(4) of type (excel)
}
