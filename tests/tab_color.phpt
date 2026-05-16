--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('tab_color.xlsx', 'Q1')
    ->setTabColor(0xFF6600)
    ->addSheet('Q2')
    ->setTabColor(\Vtiful\Kernel\Format::COLOR_GREEN)
    ->insertText(0, 0, 'x')
    ->output();

var_dump(is_file($path));

/* Round-trip: each sheet's tab colour is in its <sheetPr><tabColor rgb=.../></sheetPr>.
 * No reader API for tab colour; probe the raw OOXML instead. */
foreach (['sheet1' => 'FFFF6600', 'sheet2' => 'FF008000'] as $sheet => $expected) {
    $xml = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/worksheets/' . $sheet . '.xml');
    preg_match('/<tabColor[^>]*\srgb="([0-9A-F]+)"/i', $xml, $m);
    echo "$sheet rgb: "    . ($m[1] ?? '(none)')                 . "\n";
    echo "$sheet match: "  . var_export(($m[1] ?? '') === $expected, true) . "\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tab_color.xlsx');
?>
--EXPECT--
bool(true)
sheet1 rgb: FFFF6600
sheet1 match: true
sheet2 rgb: FF008000
sheet2 match: true
