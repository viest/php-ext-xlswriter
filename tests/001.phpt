--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
echo "xlswriter extension is available";
?>
--EXPECT--
xlswriter extension is available
