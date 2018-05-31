--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->data([['wjx', 21]]);

$filePath = $fileObject->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial01.xlsx');
?>
--EXPECT--
string(23) "./tests/tutorial01.xlsx"
