--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
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
/* Plain list() destructuring instead of `as [$range, $cf]` so the test
 * parses on PHP 7.0 (short-list destructuring is a 7.1 feature). */
foreach ($rules as $rule) {
    list($range, $cf) = $rule;
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

/* Round-trip: the 6 rules are recoverable through the reader. */
$cfs = (new \Vtiful\Kernel\Excel($config))
    ->openFile('conditional_format.xlsx')->openSheet()->getConditionalFormats();
$ranges = array_column($cfs, 'range');
sort($ranges);
echo "ranges: " . implode(',', $ranges) . "\n";
echo "ruleTypes: " . implode(',', array_map(function($cf) { return $cf['rules'][0]['type']; }, $cfs)) . "\n";

/* Round-trip the single-cell variant: TYPE_CELL/CRITERIA_EQUAL_TO/"hot" preserved. */
$cf2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('conditional_format_cell.xlsx')->openSheet()->getConditionalFormats();
echo "cellRange: " . $cf2[0]['range'] . "\n";
echo "cellRuleType: " . $cf2[0]['rules'][0]['type'] . "\n";
echo "cellOperator: " . $cf2[0]['rules'][0]['operator'] . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/conditional_format.xlsx');
@unlink(__DIR__ . '/conditional_format_cell.xlsx');
?>
--EXPECT--
bool(true)
bool(true)
ranges: B2:B10,C2:C10,D2:D10,E2:E10,F2:F10,G2:G10
ruleTypes: cellIs,containsText,aboveAverage,top10,colorScale,dataBar
cellRange: A1
cellRuleType: cellIs
cellOperator: equal
