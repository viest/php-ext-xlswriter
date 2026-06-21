--TEST--
computeFormula() makes insertFormula store the computed result (compute-on-write)
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('compute_formula.xlsx', 'S')
    ->computeFormula(true)
    ->insertText(0, 0, 10)
    ->insertText(1, 0, 20)
    ->insertText(2, 0, 30)
    ->insertFormula(3, 0, '=SUM(A1:A3)')   // -> 60
    ->insertText(0, 1, 'foo')
    ->insertFormula(1, 1, '=UPPER(B1)')    // -> FOO
    ->output();

// The saved file carries the computed cached results; read them back.
$rows = (new \Vtiful\Kernel\Excel($config))
    ->openFile('compute_formula.xlsx')
    ->openSheet('S')
    ->getSheetData();

var_dump($rows[3][0]);
var_dump($rows[1][1]);

// Confirm the cached <v> is written into the XML, not a 0 placeholder.
$xml = shell_exec('unzip -p ./tests/compute_formula.xlsx xl/worksheets/sheet1.xml');
echo 'sum cached: ', (strpos($xml, '<f>SUM(A1:A3)</f><v>60</v>') !== false ? 'yes' : 'no'), PHP_EOL;
echo 'upper cached: ', (strpos($xml, '<v>FOO</v>') !== false ? 'yes' : 'no'), PHP_EOL;
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/compute_formula.xlsx');
?>
--EXPECT--
int(60)
string(3) "FOO"
sum cached: yes
upper cached: yes
