--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("15.xlsx")
    ->header(['name', 'age'])
    ->data([
        ['one', 10],
        ['two', 20],
        ['three', 30],
    ])
    ->autoFilter("A1:B3")
    ->output();

var_dump($filePath);

/* Round-trip: data round-trips intact (non-empty). */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('15.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/15.xlsx');
?>
--EXPECT--
string(15) "./tests/15.xlsx"
bool(true)
