--TEST--
iterateImages extracts embedded image binaries
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
$png = __DIR__ . '/../resource/logo.png';
if (!is_file($png)) die('skip resource/logo.png missing');
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$png = realpath(__DIR__ . '/../resource/logo.png');
$src_size = filesize($png);

$excel->fileName('open_xlsx_iterate_images.xlsx', 'S')
    ->header(['a', 'b'])
    ->insertImage(2, 0, $png)
    ->output();

$excel->openFile('open_xlsx_iterate_images.xlsx')->openSheet('S');

$collected = [];
$excel->iterateImages(function ($img) use (&$collected) {
    $collected[] = [
        'from_row' => $img['from_row'],
        'from_col' => $img['from_col'],
        'mime'     => $img['mime'],
        'size'     => strlen($img['data']),
        'is_png'   => substr($img['data'], 0, 8) === "\x89PNG\r\n\x1a\n",
    ];
});

var_dump(count($collected));
foreach ($collected as $i) {
    printf("from=(%d,%d) mime=%s size_eq_src=%s is_png=%s\n",
        $i['from_row'], $i['from_col'], $i['mime'],
        ($i['size'] === $src_size) ? 'yes' : 'no',
        $i['is_png'] ? 'yes' : 'no');
}

/* Callback returning false should stop iteration. */
$excel->openFile('open_xlsx_iterate_images.xlsx')->openSheet('S');
$count = 0;
$excel->iterateImages(function ($img) use (&$count) {
    $count++;
    return false;
});
var_dump($count);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_iterate_images.xlsx');
?>
--EXPECT--
int(1)
from=(2,0) mime=image/png size_eq_src=yes is_png=yes
int(1)
