--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests',
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('freeze_panes.xlsx');
$fileHandle = $fileObject->getHandle();

$filePath = $fileObject->freezePanes(1, 0)
    ->header(['name', 'age']);

for ($index = 0; $index < 100; $index++) {
    $fileObject->insertText($index + 1, 0, 'wjx');
    $fileObject->insertText($index + 1, 1, 21);
}

var_dump($fileObject->output());

/* Round-trip: file opens, sheet contains the inserted rows. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('freeze_panes.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) === 101);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/freeze_panes.xlsx');
?>
--EXPECT--
string(25) "./tests/freeze_panes.xlsx"
bool(true)
