--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('005.xlsx')
    ->header(['Item', 'Cost'])
    ->output();
var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/005.xlsx');
?>
--EXPECT--
string(16) "./tests/005.xlsx"
