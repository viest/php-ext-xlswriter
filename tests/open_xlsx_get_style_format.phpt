--TEST--
getStyleFormat returns numFmt metadata for a style
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('open_xlsx_get_style_format.xlsx', 'S')
    ->header(['x', 'date'])
    ->data([[1, 0]])
    ->insertDate(1, 1, 1568389354)
    ->output();

$excel->openFile('open_xlsx_get_style_format.xlsx')->openSheet('S');

/* style 0 = General. */
$f = $excel->getStyleFormat(0);
printf("style 0: category=%s num_fmt_id=%d\n", $f['category'], $f['num_fmt_id']);

/* Probe style ids 1..5; the date column is styled with a date format. */
$found_date = false;
for ($i = 1; $i <= 8; $i++) {
    $f = $excel->getStyleFormat($i);
    if ($f === null) continue;
    if ($f['category'] === 'date' || $f['category'] === 'datetime') {
        $found_date = true;
        printf("date style: category=%s\n", $f['category']);
        break;
    }
}
echo $found_date ? "found_date=yes\n" : "found_date=no\n";

/* Out-of-range returns null. */
var_dump($excel->getStyleFormat(99999));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_style_format.xlsx');
?>
--EXPECT--
style 0: category=general num_fmt_id=0
date style: category=datetime
found_date=yes
NULL
