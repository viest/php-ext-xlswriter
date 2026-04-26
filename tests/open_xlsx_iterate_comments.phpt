--TEST--
iterateComments: writer's insertComment -> reader's iterateComments callback
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_iterate_comments.xlsx', 'S1')
    ->insertText(0, 0, 'header')
    ->insertComment(1, 0, 'Reviewed by Eve')
    ->insertComment(3, 2, 'Second note')
    ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_iterate_comments.xlsx')
    ->openSheet();

$out = [];
$r->iterateComments(function($c) use (&$out) {
    $out[] = sprintf("(%d,%d) %s", $c['row'], $c['col'], $c['text']);
});

print_r($out);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_iterate_comments.xlsx');
?>
--EXPECT--
Array
(
    [0] => (2,1) Reviewed by Eve
    [1] => (4,3) Second note
)
