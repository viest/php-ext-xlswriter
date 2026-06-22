--TEST--
TYPE_INT read of an unparsable numeric-looking string falls back to string
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* is_number() accepts shapes like "." that sscanf cannot convert. The INT path
 * used to leave the result variable uninitialized (garbage) for ".". It now
 * initializes and checks the sscanf return: a fully-unparsable value ("." )
 * falls back to a string; a partially-parsable one ("1.2.3") keeps the historic
 * leading-integer behaviour; a clean integer parses normally. */
$dir = __DIR__;
$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel->fileName('read_int_unparsable.xlsx')
    ->data([['.'], ['1.2.3'], ['42']])
    ->output();

$rows = (new \Vtiful\Kernel\Excel(['path' => $dir]))
    ->openFile('read_int_unparsable.xlsx')->openSheet()
    ->setType([\Vtiful\Kernel\Excel::TYPE_INT])
    ->getSheetData();

foreach ($rows as $r) {
    var_dump($r[0]);
}
?>
--CLEAN--
<?php @unlink(__DIR__ . '/read_int_unparsable.xlsx'); ?>
--EXPECT--
string(1) "."
int(1)
int(42)
