--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php 
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$boldStyle = \Vtiful\Kernel\Format::bold($fileHandle);

$filePath = $fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setColumn('A:A', 200, $boldStyle)
    ->output();

var_dump($filePath);
?>
--EXPECT--
string(23) "./tests/tutorial01.xlsx"