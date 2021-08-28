--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("14.xlsx")
    ->header(['name', 'money']);

for($index = 1; $index < 10; $index++) {
    $freeFile->insertText($index, 0, 'vikin');
    $freeFile->insertText($index, 1, 10);
}

$freeFile->insertText(12, 0, "Total");
$freeFile->insertFormula(12, 1, '=SUM(B2:B11)');

$filePath = $freeFile->output();

var_dump($filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/14.xlsx');
?>
--EXPECT--
string(15) "./tests/14.xlsx"
