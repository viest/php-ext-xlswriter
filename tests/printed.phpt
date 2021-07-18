--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
try {
    $config = ['path' => './tests'];
    $excel  = new \Vtiful\Kernel\Excel($config);

    $excel->setPortrait();
} catch (\Exception $exception) {
    var_dump($exception->getCode());
    var_dump($exception->getMessage());
}

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_portrait.xlsx', 'sheet1')
    ->setPortrait()
    ->output();

var_dump($excel);

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_landscape.xlsx', 'sheet1')
    ->setLandscape()
    ->output();

var_dump($excel);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/printed_portrait.xlsx');
@unlink(__DIR__ . '/printed_landscape.xlsx');
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
  string(29) "./tests/printed_portrait.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
object(Vtiful\Kernel\Excel)#1 (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(30) "./tests/printed_landscape.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
