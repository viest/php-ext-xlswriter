--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('A'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('AC'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('AB'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('AZ'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('ABC'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('ADE'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('AS'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('XF'));
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('ST'));
?>
--EXPECT--
int(0)
int(28)
int(27)
int(51)
int(730)
int(784)
int(44)
int(629)
int(513)