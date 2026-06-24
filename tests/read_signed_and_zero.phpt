--TEST--
Reader returns negative numbers and zero as numbers, not strings
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* Regression: the numeric fallback used a `> 0` guard, so 0 and every negative
 * number fell through to the string branch and came back as strings. They are
 * numbers and must round-trip as int/float. (Values beyond zend_long range are
 * still kept as strings — see open_xlsx_get_data_bignumbers.) */
$dir   = __DIR__;
$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel->fileName('read_signed_and_zero.xlsx')
    ->data([[-5], [0], [-3.14], [1000000]])
    ->output();

$data = (new \Vtiful\Kernel\Excel(['path' => $dir]))
    ->openFile('read_signed_and_zero.xlsx')->openSheet()->getSheetData();
foreach ($data as $r) {
    var_dump($r[0]);
}
?>
--CLEAN--
<?php @unlink(__DIR__ . '/read_signed_and_zero.xlsx'); ?>
--EXPECT--
int(-5)
int(0)
float(-3.14)
int(1000000)
