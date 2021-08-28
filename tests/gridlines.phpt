--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("gridlines.xlsx");

$fileObject->header(['name', 'age'])
    ->gridline(\Vtiful\Kernel\Excel::GRIDLINES_HIDE_ALL)
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
@unlink(__DIR__ . '/gridlines.xlsx');
?>
--EXPECT--
string(22) "./tests/gridlines.xlsx"
