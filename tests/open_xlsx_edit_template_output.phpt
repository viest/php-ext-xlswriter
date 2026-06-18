--TEST--
openFile can edit an existing xlsx template and output a new file
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

$writer = new \Vtiful\Kernel\Excel($config);
$writer->fileName('open_xlsx_edit_template_source.xlsx', 'Edit')
    ->insertText(0, 0, 'old')
    ->insertText(1, 1, 1)
    ->mergeCells('A4:B4', 'merged')
    ->protection('pw')
    ->output();

$out = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_source.xlsx')
    ->openSheet('Edit')
    ->insertText(0, 0, 'new text')
    ->insertText(1, 1, 42.5)
    ->insertText(1, 2, true)
    ->insertFormula(1, 3, '=B2*2')
    ->output('open_xlsx_edit_template_output.xlsx');

echo basename($out), PHP_EOL;

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_edit_template_output.xlsx')
    ->openSheet('Edit');
$rows = $reader->getSheetData();

var_dump($rows[0][0]);
var_dump($rows[1][1]);
var_dump($rows[1][2]);

$xml = shell_exec('unzip -p ./tests/open_xlsx_edit_template_output.xlsx xl/worksheets/sheet1.xml');
echo "inlineStr: " . (strpos($xml, 't="inlineStr"') !== false ? 'yes' : 'no') . PHP_EOL;
echo "formula: " . (strpos($xml, '<f>B2*2</f>') !== false ? 'yes' : 'no') . PHP_EOL;
echo "merge: " . (strpos($xml, '<mergeCells') !== false ? 'yes' : 'no') . PHP_EOL;
echo "protection: " . (strpos($xml, '<sheetProtection') !== false ? 'yes' : 'no') . PHP_EOL;

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
inlineStr: yes
formula: yes
merge: yes
protection: yes
open_xlsx_edit_template_source.xlsx
string(8) "in place"
