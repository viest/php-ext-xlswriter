<?php
/* 09 — Image round-trip: write a PNG with insertImage(),
 *      enumerate it back via iterateImages().
 */

$dir  = sys_get_temp_dir();
$name = '09_images_roundtrip.xlsx';
$png  = __DIR__ . '/../resource/pecl.png';

if (!is_file($png)) {
    fwrite(STDERR, "fixture image not found: $png\n");
    exit(1);
}

$excel = (new \Vtiful\Kernel\Excel(['path' => $dir]))
    ->fileName($name)
    ->header(['name', 'logo'])
    ->data([['xlswriter', '']])
    ->insertImage(2, 0, $png);
$path = $excel->output();

echo "wrote $path\n";

$reader = new \Vtiful\Kernel\Excel(['path' => $dir]);
$reader->openFile($name)->openSheet();

$reader->iterateImages(function (array $img) {
    printf("- mime=%s anchor=(%d,%d) size=%d bytes name=%s\n",
        $img['mime'], $img['from_row'], $img['from_col'],
        strlen($img['data']), $img['name']);
});

@unlink($path);
