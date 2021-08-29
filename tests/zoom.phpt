--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("zoom.xlsx");

$fileObject->header(['name', 'age'])
    ->zoom(100)
    ->data([
    ['viest', 21],
    ['viest', 22],
    ['viest', 23],
    ]);

$filePath = $fileObject->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/zoom.xlsx');
?>
--EXPECT--
string(17) "./tests/zoom.xlsx"
