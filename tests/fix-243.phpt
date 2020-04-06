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

$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['NumberToString', 'Number'])
    ->data([
        ['01234567', '01234567']
    ])
    ->output();

$data = $excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
    ])
    ->getSheetData();

var_dump($data);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  array(2) {
    [0]=>
    string(14) "NumberToString"
    [1]=>
    string(6) "Number"
  }
  [1]=>
  array(2) {
    [0]=>
    string(8) "01234567"
    [1]=>
    int(1234567)
  }
}