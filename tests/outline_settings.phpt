--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('outline_settings.xlsx');

$excel->outlineSettings(
    /* visible */ true,
    /* below */ false,
    /* right */ false
    /* autoStyle = false */
);

$filePath = $excel
    ->header(['Region', 'Sales'])
    ->data([
        ['North', 1000], // row 2
        ['North', 1200],
        ['North', 900],
        ['North', 1200],
        ['North Total', 4300], // row 6
        ['South', 400], // row 7
        ['South', 600],
        ['South', 500],
        ['South', 600],
        ['South Total', 2100], // row 11
        ['Grand Total', 6400], // row 12
        ['hidden row', 0],
    ])
    ->setRow('A1', 15, null)
    ->setRow('A2:A5', 15, null, 2, false, true)
    ->setRow('A6', 15, null, 1, true)
    ->setRow('A7:A10', 15, null, 2)
    ->setRow('A11', 15, null, 1)
    ->setRow('A12', 15, null, 0)
    ->setRow('A13', 15, null, null, null, true)
    ->output();

var_dump($filePath);

/* Round-trip: outlineSettings(visible=true, below=false, right=false,
 * autoStyle=false) emits <sheetPr><outlinePr summaryBelow="0"
 * summaryRight="0"/></sheetPr> in xl/worksheets/sheet1.xml (visible=true
 * is the worksheet-view default and so isn't serialised). No reader API
 * for these flags; probe the raw OOXML. */
$xml = shell_exec('unzip -p ' . escapeshellarg($filePath) . ' xl/worksheets/sheet1.xml');
preg_match('/<outlinePr([^\/>]*)\/>/', $xml, $m);
$attrs = $m[1] ?? '';
echo "outlinePr present: " . var_export(isset($m[0]), true) . "\n";
echo "summaryBelow=0: "    . var_export(strpos($attrs, 'summaryBelow="0"') !== false, true) . "\n";
echo "summaryRight=0: "    . var_export(strpos($attrs, 'summaryRight="0"') !== false, true) . "\n";
echo "no applyStyles: "    . var_export(strpos($attrs, 'applyStyles') === false, true) . "\n";

/* Flipping autoStyle adds the applyStyles attribute. */
$path2 = (new \Vtiful\Kernel\Excel($config))
    ->fileName('outline_settings_auto.xlsx')
    ->outlineSettings(true, true, true, true)
    ->insertText(0, 0, 'x')
    ->output();
$xml2 = shell_exec('unzip -p ' . escapeshellarg($path2) . ' xl/worksheets/sheet1.xml');
preg_match('/<outlinePr([^\/>]*)\/>/', $xml2, $m2);
echo "autoStyle.applyStyles=1: " . var_export(strpos($m2[1] ?? '', 'applyStyles="1"') !== false, true) . "\n";
@unlink($path2);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/outline_settings.xlsx');
@unlink(__DIR__ . '/outline_settings_auto.xlsx');
?>
--EXPECT--
string(29) "./tests/outline_settings.xlsx"
outlinePr present: true
summaryBelow=0: true
summaryRight=0: true
no applyStyles: true
autoStyle.applyStyles=1: true
