--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
var_dump(is_string(xlswriter_get_version()));
var_dump(is_string(xlswriter_get_author()));
?>
--EXPECT--
bool(true)
bool(true)
