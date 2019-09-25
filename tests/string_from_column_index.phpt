--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(0));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(28));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(27));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(51));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(730));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(784));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(44));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(629));
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(513));
?>
--EXPECT--
string(1) "A"
string(2) "AC"
string(2) "AB"
string(2) "AZ"
string(3) "ABC"
string(3) "ADE"
string(2) "AS"
string(2) "XF"
string(2) "ST"