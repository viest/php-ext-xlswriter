<?php
/* 05 — Styles: Format objects are built fluently, frozen with toResource(),
 *      and passed where a "style" parameter is accepted.
 */

$dir  = sys_get_temp_dir();
$name = '05_styles.xlsx';

$excel  = new \Vtiful\Kernel\Excel(['path' => $dir]);
$excel  = $excel->fileName($name);
$handle = $excel->getHandle();

$header = (new \Vtiful\Kernel\Format($handle))
    ->bold()
    ->fontColor(\Vtiful\Kernel\Format::COLOR_WHITE)
    ->background(\Vtiful\Kernel\Format::COLOR_BLUE)
    ->align(\Vtiful\Kernel\Format::FORMAT_ALIGN_CENTER)
    ->toResource();

$money = (new \Vtiful\Kernel\Format($handle))
    ->number('"$"#,##0.00')
    ->toResource();

$excel->header(['Name', 'Salary'])
      ->setRow('A1', 25, $header)
      ->insertText(1, 0, 'Alice')
      ->insertText(1, 1, 12345.67, null, $money)
      ->insertText(2, 0, 'Bob')
      ->insertText(2, 1,  9876.54, null, $money)
      ->output();

echo "wrote $dir/$name\n";
@unlink($dir . '/' . $name);
