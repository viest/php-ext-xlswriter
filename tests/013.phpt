--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$textFile = $excel->fileName("13.xlsx")
    ->header(['name', 'age']);

for ($index = 0; $index < 10; $index++) {
    $textFile->insertText($index+1, 0, 'vikin');
    $textFile->insertText($index+1, 1, 1000, '#,##0');
}

$filePath = $textFile->output();

var_dump($filePath);

/* Round-trip: data round-trips intact (non-empty). */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('13.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_) && count($d_) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/13.xlsx');
?>
--EXPECT--
string(15) "./tests/13.xlsx"
bool(true)
