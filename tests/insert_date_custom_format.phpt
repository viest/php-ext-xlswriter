--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('insert_date_custom_format.xlsx');

$filePath = $fileObject->header(['date'])
    ->insertDate(1, 0, time(), 'mmm d yyyy hh:mm AM/PM')
    ->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_date_custom_format.xlsx');
?>
--EXPECT--
string(38) "./tests/insert_date_custom_format.xlsx"
