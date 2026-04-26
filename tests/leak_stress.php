<?php
/*
 * Leak / memory-stability stress harness.
 *
 * Loops a writer round-trip (write → read back via reader API → close)
 * many times and asserts that PHP's reported peak memory does not grow
 * after the warm-up window. Intended to be run under:
 *
 *   USE_ZEND_ALLOC=0 valgrind --leak-check=full \
 *       --show-leak-kinds=definite,indirect --error-exitcode=1 php tests/leak_stress.php
 *
 * or on macOS:
 *
 *   USE_ZEND_ALLOC=0 MallocStackLogging=1 leaks --atExit -- \
 *       php -n -d extension=./modules/xlswriter.so tests/leak_stress.php
 *
 * USE_ZEND_ALLOC=0 is mandatory: with Zend MM pools on, all emalloc/efree
 * mismatches are masked by the per-request arena bulk-free and leak tools
 * report 0 leaks (false negative). See CLAUDE.md for the full incident.
 */

if (!extension_loaded('xlswriter')) {
    fwrite(STDERR, "xlswriter not loaded\n");
    exit(1);
}

$reader_available = method_exists(\Vtiful\Kernel\Excel::class, 'openFile');

$dir = __DIR__;
$path = $dir . '/leak_stress.xlsx';

$iterations = (int) (getenv('LEAK_STRESS_ITER') ?: 200);
$warmup     = 10;
$peak_after_warmup = 0;
$peak_final        = 0;

for ($i = 0; $i < $iterations; $i++) {
    $config = ['path' => $dir];

    /* ---- writer pass ---- */
    $excel = new \Vtiful\Kernel\Excel($config);
    $handle = $excel->fileName('leak_stress.xlsx')->getHandle();

    $fmt = (new \Vtiful\Kernel\Format($handle))
        ->fontColor(\Vtiful\Kernel\Format::COLOR_BLUE)
        ->bold()
        ->toResource();

    $excel->header(['name', 'age', 'note'])
          ->insertText(1, 0, 'viest', null, $fmt)
          ->insertText(1, 1, 21)
          ->insertText(1, 2, 'leak-stress')
          ->insertFormula(2, 1, '=SUM(B2:B2)')
          ->mergeCells('A4:C4', 'merged')
          ->output();
    unset($excel, $handle, $fmt);

    /* ---- reader pass ---- */
    if ($reader_available) {
        $rd = new \Vtiful\Kernel\Excel($config);
        $rd->openFile('leak_stress.xlsx')->openSheet();
        $rows = 0;
        while (($row = $rd->nextRow()) !== null) {
            $rows++;
        }
        if ($rows < 1) {
            fwrite(STDERR, "iter $i: reader returned no rows\n");
            exit(1);
        }
        unset($rd, $row);
    }

    if ($i === $warmup) {
        $peak_after_warmup = memory_get_peak_usage(true);
    }
}

@unlink($path);

$peak_final = memory_get_peak_usage(true);
$grew = $peak_final - $peak_after_warmup;

printf("iterations=%d  reader=%s\n", $iterations, $reader_available ? 'on' : 'off');
printf("peak_after_warmup=%d  peak_final=%d  grew=%d\n",
       $peak_after_warmup, $peak_final, $grew);

/*
 * Allow up to 256 KiB of post-warmup growth (covers Zend's lazy table
 * resizing). Anything beyond that is a real leak.
 */
$threshold = 256 * 1024;
if ($grew > $threshold) {
    fwrite(STDERR, "FAIL: peak memory grew $grew bytes (>$threshold)\n");
    exit(1);
}
echo "OK\n";
