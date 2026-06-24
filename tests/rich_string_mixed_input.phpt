--TEST--
insertRichText skips non-object array elements instead of crashing on them
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* A mixed array — RichString objects plus a plain scalar — must not treat the
 * scalar as a rich_string_object. The first pass skipped non-objects when
 * counting; the second pass used to dereference them anyway and crash. The
 * scalar is skipped and the rich fragments are written. */
$dir   = __DIR__;
$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$h     = $excel->fileName('rich_string_mixed_input.xlsx')->getHandle();

$bold   = (new \Vtiful\Kernel\Format($h))->bold()->toResource();
$italic = (new \Vtiful\Kernel\Format($h))->italic()->toResource();
$rs1 = new \Vtiful\Kernel\RichString('Hello ', $bold);
$rs2 = new \Vtiful\Kernel\RichString('world',  $italic);

$excel->insertRichText(0, 0, [$rs1, $rs2, 'ignored plain'])->output();

$data = (new \Vtiful\Kernel\Excel(['path' => $dir]))
    ->openFile('rich_string_mixed_input.xlsx')->openSheet()->getSheetData();
var_dump($data[0][0]);
echo "no crash\n";
?>
--CLEAN--
<?php @unlink(__DIR__ . '/rich_string_mixed_input.xlsx'); ?>
--EXPECT--
string(11) "Hello world"
no crash
