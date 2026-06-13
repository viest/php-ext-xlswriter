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
$excel_one  = new \Vtiful\Kernel\Excel($config);
$fileOne = $excel_one->fileName('010-1.xlsx')
    ->header(['test1'])
    ->data([['data1']])
    ->output();
$excel_two  = new \Vtiful\Kernel\Excel($config);
$fileTwo = $excel_two->fileName('010-2.xlsx')
    ->header(['test2'])
    ->data([['data2']])
    ->output();
var_dump($fileOne, $fileTwo);

/* Round-trip: each file's full sheet content is recovered byte-for-byte. */
$expected = [
    '010-1' => [['test1'], ['data1']],
    '010-2' => [['test2'], ['data2']],
];
foreach ($expected as $name => $want) {
    $got = (new \Vtiful\Kernel\Excel($config))
        ->openFile($name . '.xlsx')->openSheet()->getSheetData();
    echo "$name rows: ";
    var_dump($got);
    echo "$name match: " . var_export($got === $want, true) . "\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/010-1.xlsx');
@unlink(__DIR__ . '/010-2.xlsx');
?>
--EXPECT--
string(18) "./tests/010-1.xlsx"
string(18) "./tests/010-2.xlsx"
010-1 rows: array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(5) "test1"
  }
  [1]=>
  array(1) {
    [0]=>
    string(5) "data1"
  }
}
010-1 match: true
010-2 rows: array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(5) "test2"
  }
  [1]=>
  array(1) {
    [0]=>
    string(5) "data2"
  }
}
010-2 match: true
