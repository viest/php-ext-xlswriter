<?php
/* 11 — Read back per-cell styles.
 *
 * `nextRowWithFormula()` returns each cell as
 *   { value, type, style_id, formula? }
 * `getStyleFormat($style_id)` resolves a style_id to a structured array
 * with font / fill / border / alignment / protection / number-format.
 */

$dir  = sys_get_temp_dir();
$name = '11_read_styles.xlsx';

$writer  = (new \Vtiful\Kernel\Excel(['path' => $dir]))->fileName($name);
$handle  = $writer->getHandle();
$style   = (new \Vtiful\Kernel\Format($handle))
    ->bold()
    ->fontColor(\Vtiful\Kernel\Format::COLOR_BLUE)
    ->toResource();

$writer->header(['Name', 'Age'])
       ->setRow('A1', 25, $style)
       ->data([['viest', 21]])
       ->output();

$reader = (new \Vtiful\Kernel\Excel(['path' => $dir]))
    ->openFile($name)
    ->openSheet();

while (($row = $reader->nextRowWithFormula()) !== null) {
    foreach ($row as $cell) {
        $sid = $cell['style_id'];
        $fmt = $sid > 0 ? $reader->getStyleFormat($sid) : null;
        printf("%-12s sid=%d  font=%s\n",
            var_export($cell['value'], true),
            $sid,
            $fmt ? json_encode($fmt['font']) : 'default');
    }
}

@unlink($dir . '/' . $name);
