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

/* autoSize() must be called before the writes it should track. */
$excel = new \Vtiful\Kernel\Excel($config);
$excel->fileName('auto_size.xlsx', 'S')
    ->autoSize('A:D')
    ->insertText(0, 0, 'hi')                      // A: 2 chars
    ->insertText(0, 1, 'Hello World')              // B: 11 chars
    ->insertText(0, 2, '你好世界')                  // C: 4 CJK (display width 8)
    ->insertText(0, 3, 1234567890)                 // D: 10 digits
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

/* A 5000-char cell is clamped to Excel's 255 max width. */
$big = str_repeat('x', 5000);
$excel3 = new \Vtiful\Kernel\Excel($config);
$excel3->fileName('auto_size3.xlsx', 'S')
    ->autoSize('A:A')
    ->insertText(0, 0, $big)
    ->output();
$r3 = (new \Vtiful\Kernel\Excel($config))->openFile('auto_size3.xlsx')->openSheet();
echo "clamp:           " . (int) $r3->getColumnOptions('A')['width'] . "\n";

/* Without a range argument every column that received data is sized. */
$excel2 = new \Vtiful\Kernel\Excel($config);
$excel2->fileName('auto_size2.xlsx', 'S')
    ->autoSize()
    ->insertText(0, 0, 'short')                            // A: 5
    ->insertText(0, 5, 'a much longer piece of text here') // F: 32
    ->output();

$r2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('auto_size2.xlsx')
    ->openSheet();
echo "no-range A:      " . (int) $r2->getColumnOptions('A')['width'] . "\n";
echo "no-range F:      " . (int) $r2->getColumnOptions('F')['width'] . "\n";

/* Without autoSize() the writes are not tracked at all. */
$excel4 = new \Vtiful\Kernel\Excel($config);
$excel4->fileName('auto_size4.xlsx', 'S')
    ->insertText(0, 0, 'not tracked')
    ->output();
$r4 = (new \Vtiful\Kernel\Excel($config))->openFile('auto_size4.xlsx')->openSheet();
echo "no-autoSize A:   ";
var_dump($r4->getColumnOptions('A'));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/auto_size.xlsx');
@unlink(__DIR__ . '/auto_size2.xlsx');
@unlink(__DIR__ . '/auto_size3.xlsx');
@unlink(__DIR__ . '/auto_size4.xlsx');
?>
--EXPECT--
A (hi):          2
B (Hello World): 11
C (你好世界):     8
D (1234567890):  10
clamp:           255
no-range A:      5
no-range F:      32
no-autoSize A:   NULL
