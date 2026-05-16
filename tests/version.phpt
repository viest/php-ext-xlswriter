--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$v = xlswriter_get_version();
$a = xlswriter_get_author();
var_dump(is_string($v) && preg_match('/^\d+\.\d+\.\d+/', $v) === 1);
var_dump(is_string($a) && strlen($a) > 0);
?>
--EXPECT--
bool(true)
bool(true)
