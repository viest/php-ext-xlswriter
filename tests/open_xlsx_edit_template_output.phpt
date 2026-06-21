--TEST--
openFile can edit an existing xlsx template, preserve other sheets/shared strings, and reject unsupported edit-mode ops
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

// Source workbook: sheet "Edit" carries the cells we will rewrite plus a
// merge + protection that must survive editing; sheet "Keep" carries
// shared-string data (the duplicate "keep me" forces sharedStrings.xml)
// that must stay intact and readable after we patch a different sheet.
$writer = new \Vtiful\Kernel\Excel($config);
$writer->fileName('open_xlsx_edit_template_source.xlsx', 'Edit')
    ->insertText(0, 0, 'old')
    ->insertText(1, 1, 1)
    ->mergeCells('A4:B4', 'merged')
    ->protection('pw');
$writer->addSheet('Keep')
    ->insertText(0, 0, 'keep me')
    ->insertText(1, 0, 'keep me')
    ->insertText(2, 0, 'untouched');
$writer->output();

$out = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_source.xlsx')
    ->openSheet('Edit')
    ->insertText(0, 0, 'new text')
    ->insertText(1, 1, 42.5)
    ->insertText(1, 2, true)
    ->insertFormula(1, 3, '=B2*2')
    ->output('open_xlsx_edit_template_output.xlsx');

echo basename($out), PHP_EOL;

// Edited sheet has the new values.
$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_output.xlsx')
    ->openSheet('Edit');
$rows = $reader->getSheetData();
var_dump($rows[0][0]);
var_dump($rows[1][1]);
var_dump($rows[1][2]);

// Untouched sheet — including its shared strings — survives the edit.
$keep = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_output.xlsx')
    ->openSheet('Keep')
    ->getSheetData();
var_dump($keep[0][0]);
var_dump($keep[1][0]);
var_dump($keep[2][0]);

$xml = shell_exec('unzip -p ./tests/open_xlsx_edit_template_output.xlsx xl/worksheets/sheet1.xml');
echo "inlineStr: " . (strpos($xml, 't="inlineStr"') !== false ? 'yes' : 'no') . PHP_EOL;
echo "formula: " . (strpos($xml, '<f>B2*2</f>') !== false ? 'yes' : 'no') . PHP_EOL;
echo "merge: " . (strpos($xml, '<mergeCells') !== false ? 'yes' : 'no') . PHP_EOL;
echo "protection: " . (strpos($xml, '<sheetProtection') !== false ? 'yes' : 'no') . PHP_EOL;

// Editing in place (no new filename) overwrites the source.
$overwrite = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_source.xlsx')
    ->openSheet('Edit')
    ->insertText(0, 0, 'in place')
    ->output();
echo basename($overwrite), PHP_EOL;

$sourceRows = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_source.xlsx')
    ->openSheet('Edit')
    ->getSheetData();
var_dump($sourceRows[0][0]);

// Operations that edit mode cannot patch must fail loudly (throw) rather
// than silently drop the data. Each op gets a fresh edit session.
function reject($config, $label, callable $op) {
    $x = (new \Vtiful\Kernel\Excel($config))
        ->openFile('open_xlsx_edit_template_source.xlsx')
        ->openSheet('Edit');
    try {
        $op($x);
        echo "$label: no-throw\n";
    } catch (\Throwable $e) {
        echo "$label: throw\n";
    }
}

reject($config, 'insertUrl', function ($x) { $x->insertUrl(4, 0, 'https://example.org'); });
reject($config, 'insertComment', function ($x) { $x->insertComment(5, 0, 'note'); });
reject($config, 'insertImage', function ($x) { $x->insertImage(6, 0, __FILE__); });
reject($config, 'autoFilter', function ($x) { $x->autoFilter('A1:B2'); });
reject($config, 'freezePanes', function ($x) { $x->freezePanes(1, 0); });
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_edit_template_source.xlsx');
@unlink(__DIR__ . '/open_xlsx_edit_template_output.xlsx');
?>
--EXPECT--
open_xlsx_edit_template_output.xlsx
string(8) "new text"
float(42.5)
int(1)
string(7) "keep me"
string(7) "keep me"
string(9) "untouched"
inlineStr: yes
formula: yes
merge: yes
protection: yes
open_xlsx_edit_template_source.xlsx
string(8) "in place"
insertUrl: throw
insertComment: throw
insertImage: throw
autoFilter: throw
freezePanes: throw
