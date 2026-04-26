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
$fileObject = $fileObject->fileName('protection.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection()
    ->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('protection.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/protection.xlsx');
?>
--EXPECT--
string(23) "./tests/protection.xlsx"
bool(true)
