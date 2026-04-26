--TEST--
setBackgroundImage / setBackgroundImageBuffer
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

/* Make a tiny PNG fixture on disk for the path-based variant. */
$png = base64_decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORK5CYII='
);
$tmp = __DIR__ . '/_bg_tmp.png';
file_put_contents($tmp, $png);

(new \Vtiful\Kernel\Excel($config))
    ->fileName('set_background_image.xlsx', 'Path')
    ->setBackgroundImage($tmp)
    ->addSheet('Buf')
    ->setBackgroundImageBuffer($png)
    ->insertText(0, 0, 'x')
    ->output();

var_dump(is_file('./tests/set_background_image.xlsx'));
@unlink($tmp);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/set_background_image.xlsx');
@unlink(__DIR__ . '/_bg_tmp.png');
?>
--EXPECT--
bool(true)
