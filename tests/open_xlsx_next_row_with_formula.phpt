--TEST--
nextRowWithFormula returns rich cell metadata
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('open_xlsx_next_row_with_formula.xlsx', 'S')
    ->header(['n', 's'])
    ->data([
        [42, 'hello'],
        [3.14, 'world'],
    ])
    ->output();

$excel->openFile('open_xlsx_next_row_with_formula.xlsx')->openSheet('S');

$row = $excel->nextRowWithFormula();
ksort($row);
foreach ($row as $idx => $cell) {
    printf("[%d] type=%s value=%s style_id=%d\n",
        $idx, $cell['type'], var_export($cell['value'], true), $cell['style_id']);
}

$row = $excel->nextRowWithFormula();
ksort($row);
foreach ($row as $idx => $cell) {
    printf("[%d] type=%s value=%s\n",
        $idx, $cell['type'], var_export($cell['value'], true));
}

$row = $excel->nextRowWithFormula();
ksort($row);
foreach ($row as $idx => $cell) {
    printf("[%d] type=%s value=%s\n",
        $idx, $cell['type'], var_export($cell['value'], true));
}

var_dump($excel->nextRowWithFormula());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row_with_formula.xlsx');
?>
--EXPECT--
[0] type=string value='n' style_id=0
[1] type=string value='s' style_id=0
[0] type=number value=42
[1] type=string value='hello'
[0] type=number value=3.14
[1] type=string value='world'
NULL
