--TEST--
insertChart() in edit mode: add a chart to an existing xlsx, keep the data
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

// Source workbook with one sheet "Data" and no _rels part.
(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_chart_src.xlsx', 'Data')
    ->header(['month', 'val'])
    ->data([['a', 10], ['b', 20], ['c', 30]])
    ->output();

// Open it and insert a column chart over the data.
$excel  = new \Vtiful\Kernel\Excel($config);
$excel->openFile('edit_chart_src.xlsx');
$handle = $excel->getHandle();
$chart  = new \Vtiful\Kernel\Chart($handle, \Vtiful\Kernel\Chart::CHART_COLUMN);
$chartResource = $chart->series('Data!$B$2:$B$4')->toResource();
$out = $excel->insertChart(5, 3, $chartResource)->output('edit_chart_out.xlsx');

echo basename($out), PHP_EOL;

// Reopen: the original data still reads back (sheet + new rels are valid).
$reader = new \Vtiful\Kernel\Excel($config);
$rows = $reader->openFile('edit_chart_out.xlsx')->openSheet('Data')->getSheetData();
$flat = '';
foreach ($rows as $r) {
    $flat .= '[' . implode(',', $r) . ']';
}
echo $flat, PHP_EOL;

// The chart parses back through the reader (well-formed drawing + chart parts).
$count = 0;
$chartReader = new \Vtiful\Kernel\Excel($config);
$chartReader->openFile('edit_chart_out.xlsx')->openSheet('Data')
    ->iterateCharts(function ($c) use (&$count) {
        $count++;
    });
echo 'charts: ', $count, PHP_EOL;
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_chart_src.xlsx');
@unlink(__DIR__ . '/edit_chart_out.xlsx');
?>
--EXPECT--
edit_chart_out.xlsx
[month,val][a,10][b,20][c,30]
charts: 1
