--TEST--
Excel::getFormulaAst: tokenises and parses common Excel formulae
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
function root_kind($t) { return $t['kind']; }

/* literal — number */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=42');
printf("number: %s/%s value=%s\n",
    $t['kind'], $t['subtype'], var_export($t['value'], true));

/* literal — string */
$t = \Vtiful\Kernel\Excel::getFormulaAst('="hello"');
printf("string: %s/%s value=%s\n",
    $t['kind'], $t['subtype'], $t['value']);

/* literal — bool, error */
$t = \Vtiful\Kernel\Excel::getFormulaAst('TRUE');
printf("bool: %s/%s value=%s\n",
    $t['kind'], $t['subtype'], var_export($t['value'], true));
$t = \Vtiful\Kernel\Excel::getFormulaAst('=#DIV/0!');
printf("err: %s/%s value=%s\n",
    $t['kind'], $t['subtype'], $t['value']);

/* ref */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=A$1');
printf("ref: %s text=%s\n", $t['kind'], $t['text']);

/* range */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=A1:B2');
printf("range: %s op=%s left=%s right=%s\n",
    $t['kind'], $t['op'], $t['args'][0]['text'], $t['args'][1]['text']);

/* function */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=SUM(A1:A10)');
printf("fn: %s name=%s\n", $t['kind'], $t['name']);

/* binary precedence: + lower than * */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=1+2*3');
printf("prec: root=%s op=%s rhs.op=%s\n",
    $t['kind'], $t['op'], $t['args'][1]['op']);

/* comparison */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=A1>=10');
printf("cmp: op=%s\n", $t['op']);

/* unary minus */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=-A1');
printf("unary: op=%s arity=%s\n", $t['op'], $t['arity']);

/* postfix percent */
$t = \Vtiful\Kernel\Excel::getFormulaAst('=A1%');
printf("pct: op=%s arity=%s\n", $t['op'], $t['arity']);
?>
--EXPECT--
number: literal/number value=42.0
string: literal/string value=hello
bool: literal/bool value=true
err: literal/error value=#DIV/0!
ref: ref text=A$1
range: op op=: left=A1 right=B2
fn: fn name=SUM
prec: root=op op=+ rhs.op=*
cmp: op=>=
unary: op=- arity=unary
pct: op=% arity=postfix
