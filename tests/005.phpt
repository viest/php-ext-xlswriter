--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("excel_writer")) print "skip"; ?>
--FILE--
<?php 
echo "output path";
?>
--EXPECT--
output path
