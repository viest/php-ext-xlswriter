--TEST--
getRowOptions / getColumnOptions / getDefaultRowHeight via fixture
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase1.xlsx')
    ->openSheet();

echo "row[1] (1-based 2, custom height 30.5):\n";
var_dump($reader->getRowOptions(1));

echo "row[2] (1-based 3, hidden):\n";
var_dump($reader->getRowOptions(2));

echo "row[3] (1-based 4, outline 2):\n";
var_dump($reader->getRowOptions(3));

echo "row[4] (1-based 5, no metadata): ";
var_dump($reader->getRowOptions(4));

echo "col B (width 22.5):\n";
var_dump($reader->getColumnOptions('B'));

echo "col C (hidden):\n";
var_dump($reader->getColumnOptions('C'));

echo "col D (covered by min=3 max=4 hidden):\n";
var_dump($reader->getColumnOptions('D'));

echo "col E (outlineLevel 2):\n";
var_dump($reader->getColumnOptions('E'));

echo "col Z (none): ";
var_dump($reader->getColumnOptions('Z'));

echo "default row height: ";
var_dump($reader->getDefaultRowHeight());

echo "default col width: ";
var_dump($reader->getDefaultColumnWidth());
?>
--EXPECT--
row[1] (1-based 2, custom height 30.5):
array(5) {
  ["height"]=>
  float(30.5)
  ["hidden"]=>
  bool(false)
  ["outline_level"]=>
  int(0)
  ["collapsed"]=>
  bool(false)
  ["custom_height"]=>
  bool(true)
}
row[2] (1-based 3, hidden):
array(5) {
  ["height"]=>
  NULL
  ["hidden"]=>
  bool(true)
  ["outline_level"]=>
  int(0)
  ["collapsed"]=>
  bool(false)
  ["custom_height"]=>
  bool(false)
}
row[3] (1-based 4, outline 2):
array(5) {
  ["height"]=>
  NULL
  ["hidden"]=>
  bool(false)
  ["outline_level"]=>
  int(2)
  ["collapsed"]=>
  bool(false)
  ["custom_height"]=>
  bool(false)
}
row[4] (1-based 5, no metadata): NULL
col B (width 22.5):
array(4) {
  ["width"]=>
  float(22.5)
  ["hidden"]=>
  bool(false)
  ["outline_level"]=>
  int(0)
  ["collapsed"]=>
  bool(false)
}
col C (hidden):
array(4) {
  ["width"]=>
  float(11)
  ["hidden"]=>
  bool(true)
  ["outline_level"]=>
  int(0)
  ["collapsed"]=>
  bool(false)
}
col D (covered by min=3 max=4 hidden):
array(4) {
  ["width"]=>
  float(11)
  ["hidden"]=>
  bool(true)
  ["outline_level"]=>
  int(0)
  ["collapsed"]=>
  bool(false)
}
col E (outlineLevel 2):
array(4) {
  ["width"]=>
  float(8)
  ["hidden"]=>
  bool(false)
  ["outline_level"]=>
  int(2)
  ["collapsed"]=>
  bool(false)
}
col Z (none): NULL
default row height: float(18.5)
default col width: float(9.25)
