--TEST--
addSheet() in edit mode: append a new sheet to an existing workbook, keep the original
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

// Source workbook with a single sheet "Orig".
(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_add_sheet_src.xlsx', 'Orig')
    ->insertText(0, 0, 'name')
    ->insertText(0, 1, 'age')
    ->insertText(1, 0, 'Alice')
    ->insertText(1, 1, 30)
    ->output();

// Open it, append a new sheet "Extra" and write to it.
$out = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_add_sheet_src.xlsx')
    ->addSheet('Extra')
    ->insertText(0, 0, 'city')
    ->insertText(0, 1, 'qty')
    ->insertText(1, 0, 'NYC')
    ->insertText(1, 1, 100)
    ->output('edit_add_sheet_out.xlsx');

echo basename($out), PHP_EOL;

// Reopen: both sheets present, original first, new one appended.
$reader = new \Vtiful\Kernel\Excel($config);
$names = $reader->openFile('edit_add_sheet_out.xlsx')->sheetList();
echo 'sheets: ', implode(', ', $names), PHP_EOL;

$fmt = function ($rows) {
    $out = '';
    foreach ($rows as $r) {
        $out .= '[' . implode(',', $r) . ']';
    }
    return $out;
};
echo 'Orig: ', $fmt($reader->openSheet('Orig')->getSheetData()), PHP_EOL;
echo 'Extra: ', $fmt($reader->openSheet('Extra')->getSheetData()), PHP_EOL;
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_add_sheet_src.xlsx');
@unlink(__DIR__ . '/edit_add_sheet_out.xlsx');
?>
--EXPECT--
edit_add_sheet_out.xlsx
sheets: Orig, Extra
Orig: [name,age][Alice,30]
Extra: [city,qty][NYC,100]
