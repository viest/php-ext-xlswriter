--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('fix-207.xlsx')
    ->header(['Name', 'Code'])
    ->data([
        ['Viest', '00024']
    ])
    ->output();

$dataOne = $excel->openFile('fix-207.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_STRING,
    ])
    ->getSheetData();

$dataTwo = $excel->openFile('fix-207.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_INT,
    ])
    ->getSheetData();

var_dump($dataOne);
var_dump($dataTwo);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/fix-207.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "Name"
    [1]=>
    string(4) "Code"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "Viest"
    [1]=>
    string(5) "00024"
  }
}
array(2) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "Name"
    [1]=>
    string(4) "Code"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "Viest"
    [1]=>
    int(24)
  }
}
