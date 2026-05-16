--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('page_setup.xlsx')
    ->repeatRows('1:3')
    ->repeatColumns('A:C')
    ->printArea('A1:H100')
    ->horizontalPageBreaks([20, 40, 60])
    ->verticalPageBreaks([5, 10])
    ->fitToPages(1, 0)
    ->insertText(0, 0, 'x')
    ->output();

var_dump(is_file($path));

/* Round-trip: rowBreaks/colBreaks live in sheet1.xml; printArea/printTitles
 * live in workbook.xml as definedNames; fitToPage shows up in sheetPr. No
 * reader API; probe the raw OOXML. */
$sheet = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/worksheets/sheet1.xml');
$wb    = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/workbook.xml');
preg_match_all('/<rowBreaks[^>]*>(.*?)<\/rowBreaks>/s', $sheet, $rb);
preg_match_all('/<colBreaks[^>]*>(.*?)<\/colBreaks>/s', $sheet, $cb);
preg_match_all('/<brk\s+id="(\d+)"/', ($rb[1][0] ?? ''), $rbIds);
preg_match_all('/<brk\s+id="(\d+)"/', ($cb[1][0] ?? ''), $cbIds);
echo "rowBreaks: " . implode(',', $rbIds[1]) . "\n";
echo "colBreaks: " . implode(',', $cbIds[1]) . "\n";
echo "fitToPage: " . var_export(strpos($sheet, '<pageSetUpPr fitToPage="1"/>') !== false, true) . "\n";

preg_match('/<definedName[^>]*name="_xlnm\.Print_Area"[^>]*>([^<]+)/', $wb, $pa);
preg_match('/<definedName[^>]*name="_xlnm\.Print_Titles"[^>]*>([^<]+)/', $wb, $pt);
echo "printArea: "   . ($pa[1] ?? '') . "\n";
echo "printTitles: " . ($pt[1] ?? '') . "\n";

/* Reject malformed ranges. */
try {
    (new \Vtiful\Kernel\Excel($config))
        ->fileName('page_setup_bad.xlsx')
        ->repeatRows('not-a-range');
    echo "no exception\n";
} catch (\Exception $e) {
    echo "caught: " . $e->getCode() . "\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/page_setup.xlsx');
@unlink(__DIR__ . '/page_setup_bad.xlsx');
?>
--EXPECT--
bool(true)
rowBreaks: 20,40,60
colBreaks: 5,10
fitToPage: true
printArea: Sheet1!$A$1:$H$100
printTitles: Sheet1!$A:$C,Sheet1!$1:$3
caught: 220
