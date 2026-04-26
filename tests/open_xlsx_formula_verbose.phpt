--TEST--
nextRowWithFormula verbose mode: full {type,text,ref,si,is_dynamic,cached_value} shape
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* Default mode: formula stays a string. */
$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase3.xlsx')
    ->openSheet();
$row = $reader->nextRowWithFormula();
echo "default formula: "; var_dump($row[0]['formula']);

/* Verbose mode. */
$rv = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase3.xlsx')
    ->openSheet(null, \Vtiful\Kernel\Excel::FORMULA_VERBOSE);

echo "row 1 (array master):\n";
$row = $rv->nextRowWithFormula();
var_dump($row[0]['formula']);

echo "row 2 (skip):\n";
$rv->nextRowWithFormula();

echo "row 3 (shared master):\n";
$row = $rv->nextRowWithFormula();
var_dump($row[0]['formula']);

echo "row 4 (shared follower — no text):\n";
$row = $rv->nextRowWithFormula();
var_dump($row[0]['formula']);
?>
--EXPECT--
default formula: string(16) "TRANSPOSE(C1:C2)"
row 1 (array master):
array(6) {
  ["type"]=>
  string(5) "array"
  ["text"]=>
  string(16) "TRANSPOSE(C1:C2)"
  ["ref"]=>
  string(5) "A1:B2"
  ["si"]=>
  NULL
  ["is_dynamic"]=>
  bool(false)
  ["cached_value"]=>
  int(10)
}
row 2 (skip):
row 3 (shared master):
array(6) {
  ["type"]=>
  string(6) "shared"
  ["text"]=>
  string(4) "B3*2"
  ["ref"]=>
  string(5) "A3:A5"
  ["si"]=>
  int(0)
  ["is_dynamic"]=>
  bool(false)
  ["cached_value"]=>
  int(2)
}
row 4 (shared follower — no text):
array(6) {
  ["type"]=>
  string(6) "shared"
  ["text"]=>
  string(0) ""
  ["ref"]=>
  NULL
  ["si"]=>
  int(0)
  ["is_dynamic"]=>
  bool(false)
  ["cached_value"]=>
  int(4)
}
