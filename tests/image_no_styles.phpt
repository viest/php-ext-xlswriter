--TEST--
insertImage round-trip: PNG inserted at (3,0) is recoverable via iterateImages with the right MIME and position
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('image_no_styles.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->insertImage(3, 0, __DIR__ . '/../resource/pecl.png')
    ->output();

var_dump($filePath);

/* Round-trip: assert image MIME and anchor position. */
$reader = new \Vtiful\Kernel\Excel($config);
$reader->openFile('image_no_styles.xlsx')->openSheet();

$imgs = [];
$reader->iterateImages(function (array $img) use (&$imgs) {
    $imgs[] = $img;
});

var_dump(count($imgs));
var_dump($imgs[0]['mime']);
var_dump($imgs[0]['from_row']);
var_dump($imgs[0]['from_col']);
var_dump(strlen($imgs[0]['data']) > 0);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/image_no_styles.xlsx');
?>
--EXPECT--
string(28) "./tests/image_no_styles.xlsx"
int(1)
string(9) "image/png"
int(3)
int(0)
bool(true)
