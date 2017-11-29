--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("excel_writer")) print "skip"; ?>
--FILE--
<?php 
echo "vtiful extension is available";
?>
--EXPECT--
vtiful extension is available
