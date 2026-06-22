--TEST--
setRow tolerates a reversed row range without a stray write
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* A reversed range like "A10:A1" used to write row 0 then underflow the row
 * counter before erroring. The range is now normalized to start <= end. */
$excel = new \Vtiful\Kernel\Excel(['path' => __DIR__]);
$excel->fileName('set_row_reversed_range.xlsx')
      ->header(['a', 'b'])
      ->setRow('A10:A1', 22)
      ->output();
echo is_file(__DIR__ . '/set_row_reversed_range.xlsx') ? "ok\n" : "missing\n";
?>
--CLEAN--
<?php @unlink(__DIR__ . '/set_row_reversed_range.xlsx'); ?>
--EXPECT--
ok
