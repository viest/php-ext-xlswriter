--TEST--
Out-of-range uint8 enum/option values throw instead of silently truncating
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* A PHP long outside 0..255 used to be cast straight to uint8_t, wrapping to a
 * different valid byte and producing a corrupt rule/option. These now throw. */
$dir   = __DIR__;
$excel = new \Vtiful\Kernel\Excel(['path' => $dir]);
$h     = $excel->fileName('narrowing_range_guard.xlsx')->getHandle();

$chart = new \Vtiful\Kernel\Chart($h, 1);
try { $chart->legendSetPosition(99999); } catch (\Vtiful\Kernel\Exception $e) { echo "legend: " . $e->getCode() . "\n"; }

$table = new \Vtiful\Kernel\Table('T');
try { $table->style(-1, 0); } catch (\Vtiful\Kernel\Exception $e) { echo "style: " . $e->getCode() . "\n"; }

echo "done\n";
?>
--CLEAN--
<?php @unlink(__DIR__ . '/narrowing_range_guard.xlsx'); ?>
--EXPECT--
legend: 192
style: 192
done
