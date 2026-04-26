--TEST--
conditionalFormatCell/Range — covers all 6 main rule families
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('conditional_format.xlsx')->getHandle();

/* Format used by several rules. */
$bold = (new \Vtiful\Kernel\Format($h))->bold()->fontColor(\Vtiful\Kernel\Format::COLOR_RED)->toResource();

$rules = [
    /* 1. Cell value > 50 */
    ['B2:B10', (new \Vtiful\Kernel\ConditionalFormat())
        ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
        ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
        ->value(50)
        ->format($bold)],

    /* 2. Text contains "ERR" */
    ['C2:C10', (new \Vtiful\Kernel\ConditionalFormat())
        ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_TEXT)
        ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_TEXT_CONTAINING)
        ->valueString('ERR')],

    /* 3. Above average */
    ['D2:D10', (new \Vtiful\Kernel\ConditionalFormat())
        ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_AVERAGE)
        ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_AVERAGE_ABOVE)],

    /* 4. Top 5 */
    ['E2:E10', (new \Vtiful\Kernel\ConditionalFormat())
        ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_TOP)
        ->value(5)],

    /* 5. 3-color scale */
    ['F2:F10', (new \Vtiful\Kernel\ConditionalFormat())
        ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_3_COLOR_SCALE)
        ->minimumColor(0xF8696B)
        ->middleColor(0xFFEB84)
        ->maximumColor(0x63BE7B)],

    /* 6. Data bar */
    ['G2:G10', (new \Vtiful\Kernel\ConditionalFormat())
        ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
        ->barColor(0x638EC6)],
];

$excel->insertText(0, 0, 'cf');
foreach ($rules as [$range, $cf]) {
    $excel->conditionalFormatRange($range, $cf);
}
$path = $excel->output();
var_dump(is_file($path));

/* Single-cell variant smoke. */
$cf = (new \Vtiful\Kernel\ConditionalFormat())
    ->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
    ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_EQUAL_TO)
    ->valueString('"hot"');

$path2 = (new \Vtiful\Kernel\Excel($config))
    ->fileName('conditional_format_cell.xlsx')
    ->insertText(0, 0, 'x')
    ->conditionalFormatCell('A1', $cf)
    ->output();
var_dump(is_file($path2));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/conditional_format.xlsx');
@unlink(__DIR__ . '/conditional_format_cell.xlsx');
?>
--EXPECT--
bool(true)
bool(true)
