<?php
/* 08 — Chart: build a line chart over the data block, anchor at D1. */

$dir  = sys_get_temp_dir();
$name = '08_chart.xlsx';

$excel  = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel  = $excel->fileName($name);
$handle = $excel->getHandle();

$chart         = new \Vtiful\Kernel\Chart($handle, \Vtiful\Kernel\Chart::CHART_LINE);
$chartResource = $chart->series('Sheet1!$A$2:$A$7')->toResource();

$path = $excel->header(['Quarterly Revenue'])
    ->data([[10], [40], [50], [20], [10], [50]])
    ->insertChart(0, 3, $chartResource)
    ->output();

echo "wrote $path\n";
@unlink($path);
