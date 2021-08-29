--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel_one  = new \Vtiful\Kernel\Excel($config);
$fileOne = $excel_one->fileName('010-1.xlsx')
    ->header(['test1'])
    ->data([
        ['data1'],
    ])
    ->output();
$excel_two  = new \Vtiful\Kernel\Excel($config);
$fileTwo = $excel_two->fileName('010-2.xlsx')
    ->header(['test2'])
    ->data([
        ['data2'],
    ])
    ->output();
var_dump($fileOne,$fileTwo);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/010-1.xlsx');
@unlink(__DIR__ . '/010-2.xlsx');
?>
--EXPECT--
string(18) "./tests/010-1.xlsx"
string(18) "./tests/010-2.xlsx"
