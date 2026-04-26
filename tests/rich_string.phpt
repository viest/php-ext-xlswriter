--TEST--
insertRichText round-trip: per-run text concatenates correctly when read back as a plain string

NOTE: per-run *style* verification (font colour for "red " vs "orange")
is a Phase 4 dependency — the reader currently flattens SST `<r>` runs.
See plans/upgrade.md §8.2.2 (`Excel::nextRowRich`).
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName("rich_string.xlsx")
    ->header(['rich string']);

$fileHandle = $fileObject->getHandle();

$format1  = new \Vtiful\Kernel\Format($fileHandle);
$colorRed = $format1->fontColor(\Vtiful\Kernel\Format::COLOR_GREEN)->toResource();

$format2     = new \Vtiful\Kernel\Format($fileHandle);
$colorOrange = $format2->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$richStringOne = new \Vtiful\Kernel\RichString('red ', $colorRed);
$richStringTwo = new \Vtiful\Kernel\RichString('orange', $colorOrange);

$fileObject->insertRichText(1, 0, [
    $richStringOne,
    $richStringTwo
]);

$filePath = $fileObject->output();

$data = $fileObject->openFile('rich_string.xlsx')
    ->openSheet()
    ->getSheetData();

var_dump($filePath);
var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/rich_string.xlsx');
?>
--EXPECT--
string(24) "./tests/rich_string.xlsx"
array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(11) "rich string"
  }
  [1]=>
  array(1) {
    [0]=>
    string(10) "red orange"
  }
}

