--TEST--
insertImage with width/height: anchor cell is preserved through round-trip
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('image_width_height_styles.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->insertImage(3, 0, __DIR__ . '/../resource/pecl.png', 10, 20)
    ->output();

var_dump($filePath);

/* Round-trip: image is recoverable, MIME and anchor row/col preserved.
 * NOTE: scale factors (10, 20) are not yet exposed by iterateImages —
 * they would round-trip via Phase 4 chart/image metadata work. */
$reader = new \Vtiful\Kernel\Excel($config);
$reader->openFile('image_width_height_styles.xlsx')->openSheet();

$imgs = [];
$reader->iterateImages(function (array $img) use (&$imgs) {
    $imgs[] = $img;
});

var_dump(count($imgs));
var_dump($imgs[0]['mime']);
var_dump($imgs[0]['from_row']);
var_dump($imgs[0]['from_col']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/image_width_height_styles.xlsx');
?>
--EXPECT--
string(38) "./tests/image_width_height_styles.xlsx"
int(1)
string(9) "image/png"
int(3)
int(0)
