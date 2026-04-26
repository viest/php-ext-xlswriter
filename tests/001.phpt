--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$ce = new ReflectionClass('Vtiful\Kernel\Excel');
echo "xlswriter extension is available", PHP_EOL;
echo "Excel class loaded: ", $ce->getName(), PHP_EOL;
echo "has openFile: ", $ce->hasMethod('openFile') ? 'yes' : 'no', PHP_EOL;
echo "has output: ",   $ce->hasMethod('output')   ? 'yes' : 'no', PHP_EOL;
?>
--EXPECT--
xlswriter extension is available
Excel class loaded: Vtiful\Kernel\Excel
has openFile: yes
has output: yes
