--TEST--
Check for vtiful presence
--SKIPIF--
<?php
if (!extension_loaded("xlswriter")) print "skip";
if (!extension_loaded("zip")) print "skip zip extension required to inspect the xlsx";
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("write_boolean.xlsx")
    ->header(['name', 'age', 'active'])
    ->data([
        ['viest', 21, true],
        ['wjx',   21, false],
    ])
    ->output();

var_dump($filePath);

/* Inspect the generated XML: bool cells must be t="b" with 1/0, not the
 * silent fall-through we got before the worksheet_write_boolean call. */
$zip = new ZipArchive();
$zip->open($filePath);
$xml = $zip->getFromName('xl/worksheets/sheet1.xml');
$zip->close();

var_dump(strpos($xml, 't="b"') !== false);
preg_match('/<c r="C2"[^>]*t="b"[^>]*>\s*<v>(\d)<\/v>/', $xml, $m1);
preg_match('/<c r="C3"[^>]*t="b"[^>]*>\s*<v>(\d)<\/v>/', $xml, $m2);
var_dump($m1[1] ?? null);
var_dump($m2[1] ?? null);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/write_boolean.xlsx');
?>
--EXPECTF--
string(%d) "%swrite_boolean.xlsx"
bool(true)
string(1) "1"
string(1) "0"
