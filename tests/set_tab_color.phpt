--TEST--
Test setTabColor method
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("set_tab_color_test.xlsx", 'Sheet1')
    ->setTabColor(\Vtiful\Kernel\Excel::COLOR_RED)
    ->data([
        ['Name', 'Age'],
        ['Alice', 30],
        ['Bob', 25],
    ])
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/set_tab_color_test.xlsx');
?>
--EXPECT--
string(34) "./tests/set_tab_color_test.xlsx"
