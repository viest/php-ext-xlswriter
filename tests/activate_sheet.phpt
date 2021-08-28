--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject->fileName('activate_sheet.xlsx')
    ->header(['name', 'age'])
    ->data([
    ['viest', 21],
    ['viest', 22],
    ['viest', 23],
    ]);

$fileObject->addSheet('twoSheet')
    ->header(['name', 'age'])
    ->data([['vikin', 22]]);

var_dump($fileObject->activateSheet('twoSheet'));

$fileObject->output();
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/activate_sheet.xlsx');
?>
--EXPECT--
bool(true)
