--TEST--
insertImage() in edit mode: add an image to an existing xlsx, keep the data
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

// A minimal valid 1x1 PNG written next to the test.
$png = base64_decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
);
file_put_contents(__DIR__ . '/edit_image_pic.png', $png);

// Source workbook with one sheet "Data" and no _rels part.
(new \Vtiful\Kernel\Excel($config))
    ->fileName('edit_image_src.xlsx', 'Data')
    ->header(['name', 'val'])
    ->data([['a', 1], ['b', 2]])
    ->output();

// Open it and insert an image.
$out = (new \Vtiful\Kernel\Excel($config))
    ->openFile('edit_image_src.xlsx')
    ->insertImage(4, 0, __DIR__ . '/edit_image_pic.png')
    ->output('edit_image_out.xlsx');

echo basename($out), PHP_EOL;

// Reopen: the original data still reads back (sheet + its new rels are valid).
$reader = new \Vtiful\Kernel\Excel($config);
$rows = $reader->openFile('edit_image_out.xlsx')->openSheet('Data')->getSheetData();
$flat = '';
foreach ($rows as $r) {
    $flat .= '[' . implode(',', $r) . ']';
}
echo $flat, PHP_EOL;

// The image is present and readable back.
$count = 0;
$imgReader = new \Vtiful\Kernel\Excel($config);
$imgReader->openFile('edit_image_out.xlsx')->openSheet('Data')
    ->iterateImages(function ($image) use (&$count) {
        $count++;
    });
echo 'images: ', $count, PHP_EOL;
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/edit_image_pic.png');
@unlink(__DIR__ . '/edit_image_src.xlsx');
@unlink(__DIR__ . '/edit_image_out.xlsx');
?>
--EXPECT--
edit_image_out.xlsx
[name,val][a,1][b,2]
images: 1
