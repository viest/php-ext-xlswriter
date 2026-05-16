--TEST--
Check for vtiful presence
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

/* Round-trip: each sheet has <picture r:id="..."/> after pageMargins AND the
 * referenced media is actually bundled at xl/media/imageN.png. No reader API
 * for the worksheet background; probe the raw OOXML. */
$path = './tests/set_background_image.xlsx';
foreach (['sheet1', 'sheet2'] as $s) {
    $xml = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/worksheets/' . $s . '.xml');
    $has = (bool)preg_match('/<picture[^>]*r:id="rId\d+"\s*\/>/', $xml);
    echo "$s hasPicture: " . var_export($has, true) . "\n";
}
$mediaCount = trim(shell_exec('unzip -l ' . escapeshellarg($path) . ' 2>&1 | grep -c "xl/media/image"'));
echo "mediaCount: " . $mediaCount . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/set_background_image.xlsx');
@unlink(__DIR__ . '/_bg_tmp.png');
?>
--EXPECT--
bool(true)
sheet1 hasPicture: true
sheet2 hasPicture: true
mediaCount: 2
