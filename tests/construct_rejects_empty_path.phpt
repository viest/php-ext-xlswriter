--TEST--
Excel::__construct rejects an empty path (guards the xls_file_path OOB read)
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
/* An empty 'path' used to underflow Z_STRLEN-1 in xls_file_path() and read out
 * of bounds. __construct now rejects it up front. */
try {
    new \Vtiful\Kernel\Excel(['path' => '']);
    echo "no exception\n";
} catch (\Vtiful\Kernel\Exception $e) {
    echo "caught: " . $e->getMessage() . "\n";
}

/* A non-empty path still constructs fine. */
$excel = new \Vtiful\Kernel\Excel(['path' => __DIR__]);
echo "non-empty ok: " . ($excel instanceof \Vtiful\Kernel\Excel ? "true" : "false") . "\n";
?>
--EXPECT--
caught: Configure 'path' must be a non-empty string type
non-empty ok: true
