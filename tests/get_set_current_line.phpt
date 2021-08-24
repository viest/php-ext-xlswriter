--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject->fileName('tutorial.xlsx')
    ->header(['name', 'age'])
    ->setCurrentLine(2)
    ->data([
        ['viest', 21],
    ]);

var_dump($fileObject->getCurrentLine());

$fileObject->output();
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
int(3)
