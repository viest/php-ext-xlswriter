--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
var_dump(\Vtiful\Kernel\Excel::timestampFromDateDouble(43727.306782407));
var_dump(\Vtiful\Kernel\Excel::timestampFromDateDouble(NULL));
var_dump(\Vtiful\Kernel\Excel::timestampFromDateDouble(43727));
?>
--EXPECT--
int(1568877706)
int(0)
int(1568851200)