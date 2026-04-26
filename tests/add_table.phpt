--TEST--
addTable: builds an Excel Table with named columns and a SUM total row
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('add_table.xlsx', 'S1')->getHandle();

/* The Table API needs the cells written *first* (without a header row) */
$excel->insertText(0, 0, 'Region')
      ->insertText(0, 1, 'Sales')
      ->insertText(1, 0, 'North')
      ->insertText(1, 1, 100)
      ->insertText(2, 0, 'South')
      ->insertText(2, 1, 200)
      ->insertText(3, 0, '');

$tab = (new \Vtiful\Kernel\Table())
    ->name('SalesTable')
    ->totalRow()
    ->style(\Vtiful\Kernel\Table::STYLE_TYPE_MEDIUM, 9)
    ->columns([
        ['header' => 'Region', 'total_string'   => 'Total'],
        ['header' => 'Sales',  'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
    ]);

$path = $excel->addTable('A1:B4', $tab)->output();
var_dump(is_file($path));

/* Reader sees the data + a SUBTOTAL row from the table's total. */
$reader = (new \Vtiful\Kernel\Excel($config))->openFile('add_table.xlsx')->openSheet();
$rows = $reader->getSheetData();
echo "rowCount: " . count($rows) . "\n";
echo "header[0]: " . $rows[0][0] . "\n";
echo "header[1]: " . $rows[0][1] . "\n";
echo "totalLabel[0]: " . $rows[3][0] . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/add_table.xlsx');
?>
--EXPECT--
bool(true)
rowCount: 4
header[0]: Region
header[1]: Sales
totalLabel[0]: Total
