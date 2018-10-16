--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial01.xlsx')
    ->header(['Item', 'Cost'])
    ->output();
var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial01.xlsx');
?>
--EXPECT--
string(23) "./tests/tutorial01.xlsx"
