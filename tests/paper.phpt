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
$fileObject = $fileObject->fileName('paper.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setPaper(\Vtiful\Kernel\Excel::PAPER_A3)
    ->setLandscape()
    ->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('paper.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/paper.xlsx');
?>
--EXPECT--
string(18) "./tests/paper.xlsx"
bool(true)
