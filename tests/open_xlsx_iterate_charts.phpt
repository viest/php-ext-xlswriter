--TEST--
iterateCharts: writer's chart -> reader's iterateCharts callback (shallow metadata)
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('open_xlsx_iterate_charts.xlsx', 'Data')->getHandle();
$excel->insertText(0, 0, 'Region')
      ->insertText(0, 1, 'Sales')
      ->insertText(1, 0, 'North')
      ->insertText(1, 1, 100)
      ->insertText(2, 0, 'South')
      ->insertText(2, 1, 200);

$chart = new \Vtiful\Kernel\Chart($h, \Vtiful\Kernel\Chart::CHART_LINE);
$chart->title('Sales overview')
      ->series('=Data!$A$2:$A$3', '=Data!$B$2:$B$3');
$excel->insertChart(4, 0, $chart->toResource())
      ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_iterate_charts.xlsx')
    ->openSheet();

$count = 0;
$r->iterateCharts(function($c) use (&$count) {
    $count++;
    printf("type=%s title=%s anchor=(%d,%d)-(%d,%d) series=%d\n",
        $c['type'], $c['title'],
        $c['anchor']['from_row'], $c['anchor']['from_col'],
        $c['anchor']['to_row'],   $c['anchor']['to_col'],
        count($c['series']));
    foreach ($c['series'] as $s) {
        printf("  cat=%s val=%s\n", $s['categories'], $s['values']);
    }
});
echo "total=$count\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_iterate_charts.xlsx');
?>
--EXPECT--
type=lineChart title=Sales overview anchor=(5,1)-(19,8) series=1
  cat=Data!$B$2:$B$3 val=Data!$A$2:$A$3
total=1
