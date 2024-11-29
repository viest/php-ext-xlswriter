--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("18.xlsx");

$fileObject->header(['name', 'age', 'active'])
    ->data([['viest', 21, true]])
    ->data([['wjx', 21, false]]);

$filePath = $fileObject->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/18.xlsx');
?>
--EXPECT--
string(15) "./tests/18.xlsx"
