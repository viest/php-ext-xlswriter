--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* 1x1 PNG written to disk so header_opt can reference it by path. */
$png = base64_decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORK5CYII='
);
$tmp = __DIR__ . '/_hdr_tmp.png';
file_put_contents($tmp, $png);

(new \Vtiful\Kernel\Excel($config))
    ->fileName('set_header_footer.xlsx')
    ->setHeader('&L&\"Calibri,Bold\"&14Hello&R&P/&N')
    ->setFooter('&L&D&R&T')
    ->insertText(0, 0, 'x')
    ->output();

(new \Vtiful\Kernel\Excel($config))
    ->fileName('set_header_footer_imgs.xlsx')
    ->setHeader('&L&G', [
        'margin'     => 0.5,
        'image_left' => $tmp,
    ])
    ->setFooter('&R&G', [
        'image_right' => $tmp,
    ])
    ->insertText(0, 0, 'x')
    ->output();

var_dump(is_file('./tests/set_header_footer.xlsx'));
var_dump(is_file('./tests/set_header_footer_imgs.xlsx'));
@unlink($tmp);

/* Round-trip: header / footer strings come back via getPageSetup. */
$ps = (new \Vtiful\Kernel\Excel($config))
    ->openFile('set_header_footer.xlsx')->openSheet()->getPageSetup();
echo "odd_header: " . $ps['odd_header'] . "\n";
echo "odd_footer: " . $ps['odd_footer'] . "\n";

/* Round-trip: the image variant emits &G tokens in the header/footer strings
 * and bundles the referenced media inside the zip. */
$ps2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('set_header_footer_imgs.xlsx')->openSheet()->getPageSetup();
echo "imgs.odd_header_hasG: " . var_export(strpos((string)$ps2['odd_header'], '&G') !== false, true) . "\n";
echo "imgs.odd_footer_hasG: " . var_export(strpos((string)$ps2['odd_footer'], '&G') !== false, true) . "\n";
$mediaCount = trim(shell_exec('unzip -l ./tests/set_header_footer_imgs.xlsx 2>&1 | grep -c "xl/media/image"'));
echo "imgs.mediaCount: " . $mediaCount . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/set_header_footer.xlsx');
@unlink(__DIR__ . '/set_header_footer_imgs.xlsx');
@unlink(__DIR__ . '/_hdr_tmp.png');
?>
--EXPECT--
bool(true)
bool(true)
odd_header: &L&\"Calibri,Bold\"&14Hello&R&P/&N
odd_footer: &L&D&R&T
imgs.odd_header_hasG: true
imgs.odd_footer_hasG: true
imgs.mediaCount: 2
