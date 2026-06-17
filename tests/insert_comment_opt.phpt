--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
$config = ['path' => './tests'];

$path = (new \Vtiful\Kernel\Excel($config))
    ->fileName('insert_comment_opt.xlsx')
    ->insertText(0, 0, 'cell')
    ->insertCommentOpt(0, 0, 'Reviewed', [
        'author'    => 'QA',
        'visible'   => \Vtiful\Kernel\Excel::COMMENT_DISPLAY_VISIBLE,
        'color'     => 0xFFFFCC,
        'font_name' => 'Calibri',
        'font_size' => 12,
        'x_offset'  => 5,
        'y_offset'  => 5,
        'x_scale'   => 1.5,
        'y_scale'   => 1.5,
    ])
    ->output();

var_dump(is_file($path));
var_dump(filesize($path) > 0);

/* Round-trip: text and author come back via iterateComments; the `visible`
 * flag lives in xl/drawings/vmlDrawing1.vml as a CSS `visibility:` property
 * and isn't surfaced by the reader, so probe the raw VML for it. */
$comments = [];
(new \Vtiful\Kernel\Excel($config))
    ->openFile('insert_comment_opt.xlsx')->openSheet()
    ->iterateComments(function ($c) use (&$comments) { $comments[] = $c; });
echo "count: "  . count($comments)        . "\n";
echo "text: "   . $comments[0]['text']    . "\n";
echo "author: " . $comments[0]['author']  . "\n";

$vml = shell_exec('unzip -p ' . escapeshellarg($path) . ' xl/drawings/vmlDrawing1.vml');
preg_match('/visibility:(visible|hidden)/', $vml, $m);
echo "vmlVisibility: " . ($m[1] ?? '(none)') . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_comment_opt.xlsx');
?>
--EXPECT--
bool(true)
bool(true)
count: 1
text: Reviewed
author: QA
vmlVisibility: visible
