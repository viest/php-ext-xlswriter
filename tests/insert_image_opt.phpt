--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* Tiny 1x1 PNG fixture. */
$png = base64_decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgAAIAAAUAAeImBZsAAAAASUVORK5CYII='
);
$tmp = __DIR__ . '/_img_opt.png';
file_put_contents($tmp, $png);

/* Single insertImageOpt call exercising every option. */
$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('insert_image_opt.xlsx', 'S')
    ->insertImageOpt(2, 1, $tmp, [
        'x_offset'        => 15,
        'y_offset'        => 7,
        'x_scale'         => 0.5,
        'y_scale'         => 0.25,
        'description'     => 'alt text',
        'url'             => 'https://example.com/',
        'tip'             => 'click me',
        'object_position' => \Vtiful\Kernel\Excel::OBJECT_DONT_MOVE_DONT_SIZE,
    ])
    ->output();
@unlink($tmp);

var_dump(is_file($path));

/* Round-trip via the reader's iterateImages. */
$got = [];
(new \Vtiful\Kernel\Excel($config))
    ->openFile('insert_image_opt.xlsx')->openSheet()
    ->iterateImages(function ($img) use (&$got) { $got[] = $img; });
echo "count: "    . count($got)               . "\n";
echo "mime: "     . $got[0]['mime']           . "\n";
echo "from_row: " . $got[0]['from_row']       . "\n";
echo "from_col: " . $got[0]['from_col']       . "\n";

/* x_offset / y_offset / object_position / description / url don't come
 * back through iterateImages — probe xl/drawings/drawing1.xml directly.
 * The 15-pixel x_offset and 7-pixel y_offset are written in EMUs
 * (1 px ≈ 9525 EMU): 15*9525=142875, 7*9525=66675. The "absolute"
 * object_position serialises as editAs="absolute" on the anchor. */
$drawing = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/drawings/drawing1.xml');
$rels    = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/drawings/_rels/drawing1.xml.rels');
preg_match('#<xdr:colOff>(\d+)</xdr:colOff>#', $drawing, $cx);
preg_match('#<xdr:rowOff>(\d+)</xdr:rowOff>#', $drawing, $ry);
echo "colOff(EMU): "    . $cx[1] . "\n";
echo "rowOff(EMU): "    . $ry[1] . "\n";
echo "editAsAbsolute: " . var_export(strpos($drawing, 'editAs="absolute"') !== false, true) . "\n";
echo "hasUrl: "         . var_export(strpos($rels,    'example.com')       !== false, true) . "\n";
echo "hasDescription: " . var_export(strpos($drawing, 'alt text')          !== false, true) . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_image_opt.xlsx');
@unlink(__DIR__ . '/_img_opt.png');
?>
--EXPECT--
bool(true)
count: 1
mime: image/png
from_row: 2
from_col: 1
colOff(EMU): 142875
rowOff(EMU): 66675
editAsAbsolute: true
hasUrl: true
hasDescription: true
