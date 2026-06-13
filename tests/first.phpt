--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
try {
    $config = ['path' => './tests'];
    $excel  = new \Vtiful\Kernel\Excel($config);

    $excel->setCurrentSheetHide();
} catch (\Exception $exception) {
    var_dump($exception->getCode());
    var_dump($exception->getMessage());
}

$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('first.xlsx', 'sheet1')
    ->addSheet('sheet2')
    ->setCurrentSheetIsFirst()
    ->output();

var_dump($excel);

/* Round-trip: setCurrentSheetIsFirst sets workbookView/@firstSheet to the
 * 0-based index of the current sheet. No reader API for this; probe the raw
 * workbook.xml. */
$xml = shell_exec('unzip -p ./tests/first.xlsx xl/workbook.xml');
preg_match('/<workbookView[^>]*\sfirstSheet="(\d+)"/', $xml, $m);
echo 'firstSheet: ' . ($m[1] ?? '(none)') . "\n";
--CLEAN--
<?php
@unlink(__DIR__ . '/first.xlsx');
?>
--EXPECTF--
int(130)
string(51) "Please create a file first, use the filename method"
object(Vtiful\Kernel\Excel)#%d (3) {
  ["config":"Vtiful\Kernel\Excel":private]=>
  array(1) {
    ["path"]=>
    string(7) "./tests"
  }
  ["fileName":"Vtiful\Kernel\Excel":private]=>
  string(18) "./tests/first.xlsx"
  ["read_row_type":"Vtiful\Kernel\Excel":private]=>
  NULL
}
firstSheet: 1
