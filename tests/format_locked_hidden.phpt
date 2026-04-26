--TEST--
Format::locked() / hidden() round-trip via reader's getStyleFormat
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('format_locked_hidden.xlsx')->getHandle();

$lockFmt   = (new \Vtiful\Kernel\Format($h))->locked()->toResource();
$hiddenFmt = (new \Vtiful\Kernel\Format($h))->hidden()->toResource();
$unlockFmt = (new \Vtiful\Kernel\Format($h))->unlocked()->toResource();

$excel->insertText(0, 0, 'locked',   null, $lockFmt)
      ->insertText(1, 0, 'hidden',   null, $hiddenFmt)
      ->insertText(2, 0, 'unlocked', null, $unlockFmt)
      ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('format_locked_hidden.xlsx')
    ->openSheet();

while ($row = $reader->nextRowWithFormula()) {
    $cell = $row[0];
    $fmt  = $reader->getStyleFormat($cell['style_id']);
    printf("%-9s locked=%s hidden=%s\n",
        $cell['value'],
        var_export((bool)$fmt['protection']['locked'], true),
        var_export((bool)$fmt['protection']['hidden'], true));
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/format_locked_hidden.xlsx');
?>
--EXPECT--
locked    locked=true hidden=false
hidden    locked=true hidden=true
unlocked  locked=false hidden=false
