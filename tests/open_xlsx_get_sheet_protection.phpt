--TEST--
getSheetProtection round-trip: writer protection() -> reader returns flags
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$unprotected = (new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_get_sheet_protection_a.xlsx')
    ->insertText(0, 0, 'plain')
    ->output();

$reader_a = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_sheet_protection_a.xlsx')
    ->openSheet();
echo "no-protection: ";
var_dump($reader_a->getSheetProtection());

(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_get_sheet_protection_b.xlsx')
    ->insertText(0, 0, 'guarded')
    ->protection('secret')
    ->output();

$reader_b = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_sheet_protection_b.xlsx')
    ->openSheet();
$p = $reader_b->getSheetProtection();
printf("password_hash_len=%d sheet=%s objects=%s scenarios=%s\n",
    strlen($p['password_hash']),
    var_export($p['sheet'], true),
    var_export($p['objects'], true),
    var_export($p['scenarios'], true));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_sheet_protection_a.xlsx');
@unlink(__DIR__ . '/open_xlsx_get_sheet_protection_b.xlsx');
?>
--EXPECT--
no-protection: NULL
password_hash_len=4 sheet=true objects=true scenarios=true
