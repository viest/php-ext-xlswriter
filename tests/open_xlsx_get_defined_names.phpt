--TEST--
getDefinedNames returns workbook-level and sheet-scoped names with hidden flag
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase1.xlsx');

var_dump($reader->getDefinedNames());
?>
--EXPECT--
array(3) {
  [0]=>
  array(4) {
    ["name"]=>
    string(10) "GlobalArea"
    ["formula"]=>
    string(17) "Visible!$A$1:$B$2"
    ["scope"]=>
    NULL
    ["hidden"]=>
    bool(false)
  }
  [1]=>
  array(4) {
    ["name"]=>
    string(13) "LocalToHidden"
    ["formula"]=>
    string(11) "Hidden!$C$3"
    ["scope"]=>
    string(6) "Hidden"
    ["hidden"]=>
    bool(false)
  }
  [2]=>
  array(4) {
    ["name"]=>
    string(9) "HiddenDef"
    ["formula"]=>
    string(13) "Visible!$Z$99"
    ["scope"]=>
    NULL
    ["hidden"]=>
    bool(true)
  }
}
