--TEST--
getConditionalFormats: writer's conditionalFormatRange -> reader's getConditionalFormats
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('open_xlsx_get_conditional_formats.xlsx')->getHandle();
$bold = (new \Vtiful\Kernel\Format($h))->bold()->toResource();

$cf1 = (new \Vtiful\Kernel\ConditionalFormat())
    ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
    ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
    ->value(50)
    ->format($bold);

$cf2 = (new \Vtiful\Kernel\ConditionalFormat())
    ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_TEXT)
    ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_TEXT_CONTAINING)
    ->valueString('ERR');

$excel->insertText(0, 0, 'x')
      ->conditionalFormatRange('A2:A10', $cf1)
      ->conditionalFormatRange('B2:B10', $cf2)
      ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_conditional_formats.xlsx')
    ->openSheet();

foreach ($r->getConditionalFormats() as $blk) {
    printf("range=%s rules=%d\n", $blk['range'], count($blk['rules']));
    foreach ($blk['rules'] as $rule) {
        printf("  type=%s op=%s formula1=%s\n",
            $rule['type'], $rule['operator'] ?? 'null', $rule['formula1'] ?? 'null');
    }
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_conditional_formats.xlsx');
?>
--EXPECT--
range=A2:A10 rules=1
  type=cellIs op=greaterThan formula1=50
range=B2:B10 rules=1
  type=containsText op=containsText formula1=NOT(ISERROR(SEARCH("ERR",B2)))
