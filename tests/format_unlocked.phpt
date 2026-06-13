--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = [
    'path' => './tests'
];

$fileObject  = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('format_unlocked.xlsx');
$fileHandle = $fileObject->getHandle();

$format       = new \Vtiful\Kernel\Format($fileHandle);
$unlockedStyle = $format->unlocked()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['wjx',   21]
    ])
    ->setRow('A2', 50, $unlockedStyle)
    ->protection()
    ->output();

var_dump($filePath);

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_unlocked.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['protection']['locked']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_unlocked.xlsx');
?>
--EXPECT--
string(28) "./tests/format_unlocked.xlsx"
bool(false)
