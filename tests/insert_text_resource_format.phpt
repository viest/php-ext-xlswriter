--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$textFile = $excel->fileName("insert_text_resource_format.xlsx")
    ->header(['name', 'age']);

$fileHandle = $textFile->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

for ($index = 0; $index < 10; $index++) {
    $textFile->insertText($index+1, 0, 'vikin');
    $textFile->insertText($index+1, 1, 1000, '#,##0', $colorStyle);
}

$filePath = $textFile->output();

var_dump($filePath);

/* Round-trip: written cells have a non-default style after read-back. */
$v_     = new \Vtiful\Kernel\Excel($config);
$v_->openFile('insert_text_resource_format.xlsx')->openSheet();
$styled_ = false;
while (($r_ = $v_->nextRowWithFormula()) !== null) {
    foreach ($r_ as $c_) { if ($c_['style_id'] > 0) { $styled_ = true; break 2; } }
}
var_dump($styled_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_text_resource_format.xlsx');
?>
--EXPECT--
string(40) "./tests/insert_text_resource_format.xlsx"
bool(true)
