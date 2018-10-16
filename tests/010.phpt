--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel_one  = new \Vtiful\Kernel\Excel($config);
$fileOne = $excel_one->fileName('tutorial01.xlsx')
    ->header(['test1'])
    ->data([
        ['data1'],
    ])
    ->output();
$excel_two  = new \Vtiful\Kernel\Excel($config);
$fileTwo = $excel_two->fileName('tutorial02.xlsx')
    ->header(['test2'])
    ->data([
        ['data2'],
    ])
    ->output();
var_dump($fileOne,$fileTwo);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial01.xlsx');
@unlink(__DIR__ . '/tutorial02.xlsx');
?>
--EXPECT--
string(23) "./tests/tutorial01.xlsx"
string(23) "./tests/tutorial02.xlsx"
