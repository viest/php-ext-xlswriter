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

/* Round-trip: each file is independently readable and holds its own header+data. */
foreach (['010-1' => ['test1', 'data1'], '010-2' => ['test2', 'data2']] as $name => $expected) {
    $rows = (new \Vtiful\Kernel\Excel($config))
        ->openFile($name . '.xlsx')->openSheet()->getSheetData();
    echo "$name header: " . $rows[0][0] . "\n";
    echo "$name data:   " . $rows[1][0] . "\n";
    echo "$name match:  " . var_export($rows[0][0] === $expected[0] && $rows[1][0] === $expected[1], true) . "\n";
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
010-1 header: test1
010-1 data:   data1
010-1 match:  true
010-2 header: test2
010-2 data:   data2
010-2 match:  true
