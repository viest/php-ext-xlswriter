--TEST--
Passing a foreign resource to Chart/Format throws instead of crashing
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* zval_get_resource() throws on a non-xlswriter resource but still returns
 * NULL; the Chart/Format constructors used to dereference that NULL. They must
 * bail on the thrown exception instead of crashing. */
$dir   = __DIR__;
$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel->fileName('foreign_resource_guard.xlsx');

$fh = fopen('php://memory', 'r+');
try {
    new \Vtiful\Kernel\Chart($fh, 1);
} catch (\Vtiful\Kernel\Exception $e) {
    echo "chart: " . $e->getCode() . "\n";
}
try {
    new \Vtiful\Kernel\Format($fh);
} catch (\Vtiful\Kernel\Exception $e) {
    echo "format: " . $e->getCode() . "\n";
}
fclose($fh);
echo "no crash\n";
?>
--CLEAN--
<?php @unlink(__DIR__ . '/foreign_resource_guard.xlsx'); ?>
--EXPECT--
chart: 210
format: 210
no crash
