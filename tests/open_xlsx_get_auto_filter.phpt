--TEST--
getAutoFilter: writer.autoFilter range round-trip + fixture column criteria
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

/* Round-trip: writer's autoFilter only sets the range — column criteria
 * aren't yet exposed in the writer API, so columns is an empty list. */
(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_get_auto_filter.xlsx')
    ->insertText(0, 0, 'x')
    ->autoFilter('A1:F10')
    ->output();

$reader = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_auto_filter.xlsx')
    ->openSheet();
$af = $reader->getAutoFilter();
printf("range=%s columns=%d\n", $af['range'], count($af['columns']));

/* No autoFilter -> NULL. */
(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_get_auto_filter_none.xlsx')
    ->insertText(0, 0, 'plain')
    ->output();
$r2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_get_auto_filter_none.xlsx')
    ->openSheet();
var_dump($r2->getAutoFilter());

/* Fixture exercises filter-list and custom-filters paths. */
$reader3 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('xlsx/phase3.xlsx')
    ->openSheet();
$af3 = $reader3->getAutoFilter();
printf("fixture range=%s\n", $af3['range']);
foreach ($af3['columns'] as $c) {
    printf("  col_id=%d type=%s", $c['col_id'], $c['type']);
    if ($c['type'] === 'list') {
        printf(" values=%s", implode(',', $c['values']));
    } elseif ($c['type'] === 'custom') {
        printf(" and=%s op1=%s val1=%s op2=%s val2=%s",
            var_export($c['and'], true),
            $c['criterion1']['operator'], $c['criterion1']['value'],
            $c['criterion2']['operator'], $c['criterion2']['value']);
    }
    echo "\n";
}
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_get_auto_filter.xlsx');
@unlink(__DIR__ . '/open_xlsx_get_auto_filter_none.xlsx');
?>
--EXPECT--
range=A1:F10 columns=0
NULL
fixture range=A1:D8
  col_id=0 type=list values=A,B
  col_id=3 type=custom and=true op1=greaterThan val1=20 op2=lessThan val2=60
