--TEST--
setHeader / setFooter (with image options)
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
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
