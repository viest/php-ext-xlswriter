--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
/* #552 / #548 / #544: object_format() used ZEND_STRL(format->val) as the
 * cache key. zend_string's `val` is a flexible array declared `char[1]`,
 * so sizeof - 1 was 0, every distinct format collapsed into the same
 * empty-string slot, and the first format wins for every subsequent
 * insertDate. Now each format is cached by its own ZSTR_VAL/ZSTR_LEN. */
$config = ['path' => './tests'];
$ts = strtotime('2026-05-19 12:34:56');

(new \Vtiful\Kernel\Excel($config))
    ->fileName('insert_date_format_cache.xlsx', 'S1')
    ->insertDate(0, 0, $ts, 'yyyy-mm-dd')
    ->insertDate(1, 0, $ts, 'hh:mm:ss')
    ->insertDate(2, 0, $ts, 'mm/dd/yy')
    ->insertDate(3, 0, $ts) /* default: yyyy-mm-dd hh:mm:ss */
    ->insertDate(4, 0, $ts, 'yyyy-mm-dd') /* same as row 0 — must reuse */
    ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('insert_date_format_cache.xlsx')->openSheet();
$row_fmts = [];
$row_sids = [];
while (($row = $r->nextRowWithFormula()) !== null) {
    $sid = $row[0]['style_id'];
    $row_sids[] = $sid;
    $row_fmts[] = $r->getStyleFormat($sid)['format_string'];
}
foreach ($row_fmts as $i => $f) {
    echo "row $i: $f\n";
}
echo "row0_sid_equals_row4_sid: " . var_export($row_sids[0] === $row_sids[4], true) . "\n";
echo "all_sids_unique_except_row0_row4: " . var_export(
    count(array_unique([$row_sids[0], $row_sids[1], $row_sids[2], $row_sids[3]])) === 4, true
) . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_date_format_cache.xlsx');
?>
--EXPECT--
row 0: yyyy-mm-dd
row 1: hh:mm:ss
row 2: mm/dd/yy
row 3: yyyy-mm-dd hh:mm:ss
row 4: yyyy-mm-dd
row0_sid_equals_row4_sid: true
all_sids_unique_except_row0_row4: true
