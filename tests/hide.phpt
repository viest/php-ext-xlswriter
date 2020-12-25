--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
try {
    $config = ['path' => './tests'];
    $excel  = new \Vtiful\Kernel\Excel($config);

    $excel->setCurrentSheetHide();
} catch (\Exception $exception) {
    var_dump($exception->getCode());
    var_dump($exception->getMessage());
}

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('hide.xlsx', 'sheet1')
    ->addSheet('sheet2')
    ->setCurrentSheetHide()
    ->output();

var_dump($excel);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/hide.xlsx');
?>
--EXPECT--
int(130)
string(51) "Please create a file first, use the filename method"
object(Vtiful\Kernel\Excel)#3 (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(17) "./tests/hide.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
