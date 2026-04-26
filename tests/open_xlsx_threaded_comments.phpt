--TEST--
iterateComments: threaded comments (with author GUIDs and parent linkage)
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase4.xlsx')
    ->openSheet();

$out = [];
$r->iterateComments(function($c) use (&$out) {
    $out[] = $c;
});

foreach ($out as $c) {
    printf("(%d,%d) thr=%s author=%s parent=%s text=%s\n",
        $c['row'], $c['col'],
        $c['threaded'] ? 'true' : 'false',
        $c['author'] ?? 'null',
        $c['parent_id'] ?? 'null',
        $c['text']);
}
?>
--EXPECT--
(1,1) thr=false author=Eve parent=null text=From Eve, visible
(2,3) thr=false author=Bob parent=null text=From Bob, hidden
(3,1) thr=true author={p1} parent=null text=Top-level threaded comment
(3,1) thr=true author={p2} parent={aaaaaaaa} text=Reply to it
