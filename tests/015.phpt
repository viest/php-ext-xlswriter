--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("15.xlsx")
    ->header(['name', 'age'])
    ->data([
        ['one', 10],
        ['two', 20],
        ['three', 30],
    ])
    ->autoFilter("A1:B3")
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/15.xlsx');
?>
--EXPECT--
string(15) "./tests/15.xlsx"
