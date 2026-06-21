--TEST--
mergeCells() works in edit mode: new merges are added, existing ones preserved
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

// Source file already has a merged range A1:C1.
(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_merge_src.xlsx', 'S')
    ->insertText(0, 0, 'title')
    ->mergeCells('A1:C1', 'title')
    ->insertText(2, 0, 1)
    ->output();

// Open it and add another merge while editing.
$out = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_merge_src.xlsx')
    ->openSheet('S')
    ->insertText(4, 0, 'merged')
    ->mergeCells('A5:C5', 'merged')
    ->output('edit_merge_out.xlsx');

echo basename($out), PHP_EOL;

// Both merges are present (getMergedCells is 1-based).
$merges = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_merge_out.xlsx')
    ->openSheet('S')
    ->getMergedCells();

foreach ($merges as $m) {
    echo "{$m['first_row']},{$m['first_col']},{$m['last_row']},{$m['last_col']}", PHP_EOL;
}

$xml = shell_exec('unzip -p ./tests/edit_merge_out.xlsx xl/worksheets/sheet1.xml');
echo 'count 2: ', (strpos($xml, '<mergeCells count="2">') !== false ? 'yes' : 'no'), PHP_EOL;
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_merge_src.xlsx');
@unlink(__DIR__ . '/edit_merge_out.xlsx');
?>
--EXPECT--
edit_merge_out.xlsx
1,1,1,3
5,1,5,3
count 2: yes
