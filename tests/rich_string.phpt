--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
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

