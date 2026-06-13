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

$fileObject = $fileObject->fileName('format_number.xlsx');
$fileHandle = $fileObject->getHandle();

$format      = new \Vtiful\Kernel\Format($fileHandle);
$numberStyle = $format->number('#,##0')->toResource();

$filePath = $fileObject->header(['name', 'balance'])
    ->data([
        ['viest', 10000],
        ['wjx',   100000]
    ])
    ->setColumn('B:B', 50, $numberStyle)
    ->output();

var_dump($filePath);

/* Round-trip: the written style records exact format values. */
$verify = new \Vtiful\Kernel\Excel($config);
$verify->openFile('format_number.xlsx')->openSheet();
$sid = 0;
while (($r = $verify->nextRowWithFormula()) !== null) {
    foreach ($r as $cell) { if ($cell['style_id'] > 0) { $sid = $cell['style_id']; break 2; } }
}
$fmt = $verify->getStyleFormat($sid);
var_dump($fmt['category'], $fmt['num_fmt_id'] >= 164);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_number.xlsx');
?>
--EXPECT--
string(26) "./tests/format_number.xlsx"
string(6) "number"
bool(true)
