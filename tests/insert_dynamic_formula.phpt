--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$file = new \Vtiful\Kernel\Excel($config);

$file->fileName('insert_dynamic_formula.xlsx', 'S')
    ->header(['A', 'B'])
    ->data([
        ['x', 1],
        ['y', 2],
        ['x', 3],
    ])
    ->insertDynamicFormula(4, 0, '=_xlfn.UNIQUE(A1:A4)')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('insert_dynamic_formula.xlsx')
    ->openSheet('S');

$found = false;
while ($row = $reader->nextRowWithFormula()) {
    foreach ($row as $col => $cell) {
        if ($cell['type'] === 'formula') {
            printf("formula=%s value=%d\n", $cell['formula'], $cell['value']);
            $found = true;
        }
    }
}
if (!$found) {
    echo "NO FORMULA FOUND\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_dynamic_formula.xlsx');
?>
--EXPECTF--
formula=_xlfn.UNIQUE(A1:A4) value=0
