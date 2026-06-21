--TEST--
evaluateFormula() computes Excel formulas, resolving refs against written cells
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$excel = new \Vtiful\Kernel\Excel(['path' => './tests']);
$excel->fileName('evaluate_formula.xlsx', 'S')
    ->insertText(0, 0, 10)        // A1
    ->insertText(1, 0, 20)        // A2
    ->insertText(2, 0, 30)        // A3
    ->insertText(0, 1, 'hello');  // B1

// Pure expressions
var_dump($excel->evaluateFormula('1+2*3'));
var_dump($excel->evaluateFormula('(1+2)*3'));
var_dump($excel->evaluateFormula('2^10'));
var_dump($excel->evaluateFormula('7/2'));

// Reference resolution against the cells written above
var_dump($excel->evaluateFormula('=SUM(A1:A3)'));
var_dump($excel->evaluateFormula('AVERAGE(A1:A3)'));
var_dump($excel->evaluateFormula('MAX(A1:A3)'));
var_dump($excel->evaluateFormula('A1+A2+A3'));

// Logical / text
var_dump($excel->evaluateFormula('IF(A1>5,"big","small")'));
var_dump($excel->evaluateFormula('CONCATENATE(B1," world")'));
var_dump($excel->evaluateFormula('UPPER(B1)'));
var_dump($excel->evaluateFormula('LEN(B1)'));

// Excel error values surface as strings
var_dump($excel->evaluateFormula('10/0'));
var_dump($excel->evaluateFormula('SQRT(-1)'));
var_dump($excel->evaluateFormula('NOPE(1)'));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/evaluate_formula.xlsx');
?>
--EXPECT--
int(7)
int(9)
int(1024)
float(3.5)
int(60)
int(20)
int(30)
int(60)
string(3) "big"
string(11) "hello world"
string(5) "HELLO"
int(5)
string(7) "#DIV/0!"
string(5) "#NUM!"
string(6) "#NAME?"
