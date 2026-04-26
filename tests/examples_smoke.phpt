--TEST--
Every script under examples/ runs without errors and exits cleanly
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$root  = realpath(__DIR__ . '/..');
$files = glob($root . '/examples/[0-9]*_*.php');
sort($files);

if (empty($files)) {
    echo "no examples found\n";
    exit(1);
}

/* Run each example in an isolated scope (closure) so they don't bleed
 * variables into each other, while still keeping the loaded extension. */
$run = static function (string $file): void {
    /* The example writes to stdout — capture it so the smoke test only
     * cares about whether the script ran without a fatal. */
    ob_start();
    try {
        require $file;
    } finally {
        ob_end_clean();
    }
};

foreach ($files as $file) {
    $name = basename($file);
    try {
        $run($file);
        echo "OK   $name\n";
    } catch (\Throwable $e) {
        printf("FAIL %s: %s\n", $name, $e->getMessage());
        exit(1);
    }
}
?>
--EXPECT--
OK   01_basic_write.php
OK   02_read_full.php
OK   03_read_streaming.php
OK   04_const_memory.php
OK   05_styles.php
OK   06_merge_cells.php
OK   07_csv_export.php
OK   08_chart.php
OK   09_images_roundtrip.php
OK   10_data_validation.php
OK   11_read_styles.php
