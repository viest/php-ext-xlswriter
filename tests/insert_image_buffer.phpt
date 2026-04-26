--TEST--
insertImageBuffer: round-trip a 1x1 PNG via in-memory buffer
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* 1x1 transparent PNG. */
$png = base64_decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORK5CYII='
);

(new \Vtiful\Kernel\Excel($config))
    ->fileName('insert_image_buffer.xlsx')
    ->insertText(0, 0, 'with image')
    ->insertImageBuffer(2, 0, $png, ['x_scale' => 2.0, 'y_scale' => 2.0])
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('insert_image_buffer.xlsx');

$count = 0;
$reader->iterateImages(function($img) use (&$count) {
    $count++;
}, 'Sheet1');

var_dump($count);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_image_buffer.xlsx');
?>
--EXPECT--
int(1)
