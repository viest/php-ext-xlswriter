--TEST--
Test addDefinedName method
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("add_defined_name_test.xlsx")
    ->addDefinedName('MyRange', '=Sheet1!$A$1:$A$10')
    ->data([
        [1],
        [2],
        [3],
    ])
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/add_defined_name_test.xlsx');
?>
--EXPECT--
string(36) "./tests/add_defined_name_test.xlsx"
