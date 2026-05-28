--TEST--
PHP bool values are written via worksheet_write_boolean (xlsx cell type "b")
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
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
$xml = "";
$zip = new ZipArchive();
if ($zip->open($filePath) === true) {
    $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
    $zip->close();
}
var_dump(strpos($xml, 't="b"') !== false);
preg_match_all('/<c r="C2"[^>]*t="b"[^>]*>\s*<v>(\d)<\/v>/', $xml, $m1);
preg_match_all('/<c r="C3"[^>]*t="b"[^>]*>\s*<v>(\d)<\/v>/', $xml, $m2);
var_dump($m1[1][0] ?? null);
var_dump($m2[1][0] ?? null);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/write_boolean.xlsx');
?>
--EXPECT--
string(26) "./tests/write_boolean.xlsx"
bool(true)
string(1) "1"
string(1) "0"
