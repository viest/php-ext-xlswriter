--TEST--
Print/page setup: repeatRows/repeatColumns/printArea/horizontalPageBreaks/verticalPageBreaks/fitToPages
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('page_setup.xlsx')
    ->repeatRows('1:3')
    ->repeatColumns('A:C')
    ->printArea('A1:H100')
    ->horizontalPageBreaks([20, 40, 60])
    ->verticalPageBreaks([5, 10])
    ->fitToPages(1, 0)
    ->insertText(0, 0, 'x')
    ->output();

var_dump(is_file($path));

/* Reject malformed ranges. */
try {
    (new \Vtiful\Kernel\Excel($config))
        ->fileName('page_setup_bad.xlsx')
        ->repeatRows('not-a-range');
    echo "no exception\n";
} catch (\Exception $e) {
    echo "caught: " . $e->getCode() . "\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/page_setup.xlsx');
@unlink(__DIR__ . '/page_setup_bad.xlsx');
?>
--EXPECT--
bool(true)
caught: 220
