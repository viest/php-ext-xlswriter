--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('close.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

$excel->close();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/close.xlsx');
?>
--EXPECT--
string(18) "./tests/close.xlsx"
