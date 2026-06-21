--TEST--
insertDate() works in edit mode: value written and a date number-format injected
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_date_src.xlsx', 'S')
    ->insertText(0, 0, 'date:')
    ->output();

$out = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_date_src.xlsx')
    ->openSheet('S')
    ->insertDate(0, 1, 1700000000, 'yyyy-mm-dd')
    ->output('edit_date_out.xlsx');

echo basename($out), PHP_EOL;

$rows = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_date_out.xlsx')
    ->openSheet('S')
    ->getSheetData();

// 1700000000 -> Excel serial 45244.92592...
echo 'value ok: ', (abs($rows[0][1] - 45244.9259259) < 0.001 ? 'yes' : 'no'), PHP_EOL;

$sheet  = shell_exec('unzip -p ./tests/edit_date_out.xlsx xl/worksheets/sheet1.xml');
$styles = shell_exec('unzip -p ./tests/edit_date_out.xlsx xl/styles.xml');
echo 'cell styled: ', (preg_match('/<c r="B1" s="\d+"/', $sheet) ? 'yes' : 'no'), PHP_EOL;
echo 'date format injected: ', (strpos($styles, 'formatCode="yyyy-mm-dd"') !== false ? 'yes' : 'no'), PHP_EOL;
echo 'applyNumberFormat: ', (strpos($styles, 'applyNumberFormat="1"') !== false ? 'yes' : 'no'), PHP_EOL;
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_date_src.xlsx');
@unlink(__DIR__ . '/edit_date_out.xlsx');
?>
--EXPECT--
edit_date_out.xlsx
value ok: yes
cell styled: yes
date format injected: yes
applyNumberFormat: yes
