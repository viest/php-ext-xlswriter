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

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$filePath = $fileObject->freezePanes(1, 0)
    ->header(['name', 'age']);

for ($index = 0; $index < 100; $index++) {
    $fileObject->insertText($index + 1, 0, 'wjx');
    $fileObject->insertText($index + 1, 1, 21);
}

var_dump($fileObject->output());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
string(21) "./tests/tutorial.xlsx"
