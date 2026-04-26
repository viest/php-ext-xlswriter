--TEST--
nextRowRich: SST rich-text runs round-trip via writer's RichString
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('open_xlsx_next_row_rich.xlsx')->getHandle();

$bold   = (new \Vtiful\Kernel\Format($h))->bold()->fontColor(\Vtiful\Kernel\Format::COLOR_RED)->toResource();
$italic = (new \Vtiful\Kernel\Format($h))->italic()->fontSize(14)->toResource();

$rs1 = new \Vtiful\Kernel\RichString('Hello ', $bold);
$rs2 = new \Vtiful\Kernel\RichString('world',  $italic);

$excel->insertRichText(0, 0, [$rs1, $rs2])
      ->insertText(1, 0, 'plain')
      ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_next_row_rich.xlsx')
    ->openSheet();

echo "rich row:\n";
$row = $r->nextRowRich();
foreach ($row[0] as $i => $run) {
    printf("  run[%d] text='%s' bold=%s italic=%s color=%s size=%s\n",
        $i, $run['text'],
        $run['font']['bold']   ? 'true' : 'false',
        $run['font']['italic'] ? 'true' : 'false',
        $run['font']['color'] ?? 'null',
        $run['font']['size']);
}

echo "plain row:\n";
$row = $r->nextRowRich();
echo "  fallback: " . $row[0][0]['text'] . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row_rich.xlsx');
?>
--EXPECT--
rich row:
  run[0] text='Hello ' bold=true italic=false color=FFFF0000 size=11
  run[1] text='world' bold=false italic=true color=null size=14
plain row:
  fallback: plain
