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
$fileObject = $fileObject->fileName('insert_date_resource_format.xlsx');

$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$filePath = $fileObject->header(['date'])
    ->insertDate(1, 0, time(), 'mmm d yyyy hh:mm AM/PM', $colorStyle)
    ->output();

var_dump($filePath);

/* Round-trip: written cells have a non-default style after read-back. */
$v_     = new \Vtiful\Kernel\Excel($config);
$v_->openFile('insert_date_resource_format.xlsx')->openSheet();
$styled_ = false;
while (($r_ = $v_->nextRowWithFormula()) !== null) {
    foreach ($r_ as $c_) { if ($c_['style_id'] > 0) { $styled_ = true; break 2; } }
}
var_dump($styled_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_date_resource_format.xlsx');
?>
--EXPECT--
string(40) "./tests/insert_date_resource_format.xlsx"
bool(true)
