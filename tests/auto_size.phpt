--TEST--
Excel::autoSize() sizes columns to content width
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$excel->fileName('auto_size.xlsx', 'S')
    ->insertText(0, 0, 'hi')                      // A: 2 chars
    ->insertText(0, 1, 'Hello World')              // B: 11 chars
    ->insertText(0, 2, '你好世界')                  // C: 4 CJK (display width 8)
    ->insertText(0, 3, 1234567890)                 // D: 10 digits
    ->autoSize('A:D')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('auto_size.xlsx')
    ->openSheet();

/* libxlsxwriter adds Excel's standard column margin (~0.71) on top of the
 * content width, so the integer part equals the content display width. */
echo "A (hi):          " . (int) $reader->getColumnOptions('A')['width'] . "\n";
echo "B (Hello World): " . (int) $reader->getColumnOptions('B')['width'] . "\n";
echo "C (你好世界):     " . (int) $reader->getColumnOptions('C')['width'] . "\n";
echo "D (1234567890):  " . (int) $reader->getColumnOptions('D')['width'] . "\n";

/* Without a range argument every column that received data is sized. */
$excel2 = new \Vtiful\Kernel\Excel($config);
$excel2->fileName('auto_size2.xlsx', 'S')
    ->insertText(0, 0, 'short')                            // A: 5
    ->insertText(0, 5, 'a much longer piece of text here') // F: 32
    ->autoSize()
    ->output();

$r2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('auto_size2.xlsx')
    ->openSheet();
echo "no-range A:      " . (int) $r2->getColumnOptions('A')['width'] . "\n";
echo "no-range F:      " . (int) $r2->getColumnOptions('F')['width'] . "\n";
echo "no-range E (no content): ";
var_dump($r2->getColumnOptions('E'));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/auto_size.xlsx');
@unlink(__DIR__ . '/auto_size2.xlsx');
?>
--EXPECT--
A (hi):          2
B (Hello World): 11
C (你好世界):     8
D (1234567890):  10
no-range A:      5
no-range F:      32
no-range E (no content): NULL
