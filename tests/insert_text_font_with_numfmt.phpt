--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
?>
--FILE--
<?php
/* #545 / #472: insertText() with both a num-format string and a Format
 * resource dropped the resource's font(), because format_copy() never
 * copied the font_name (and font_scheme / has_font / has_dxf_font)
 * fields. Every clone fell back to the workbook default (Calibri). */
$config = ['path' => './tests'];

$excel = new \Vtiful\Kernel\Excel($config);
$h = $excel->fileName('insert_text_font_with_numfmt.xlsx', 'S')->getHandle();
$style = (new \Vtiful\Kernel\Format($h))->font('Arial')->bold()->toResource();

$excel
    ->insertText(0, 0, 1000, null,         $style)
    ->insertText(0, 1, 1000, '#,##0.00',   $style)
    ->insertText(1, 0, 1000, '#,##0.0000', $style)
    ->output();

$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('insert_text_font_with_numfmt.xlsx')->openSheet();

$sids = [];
while (($row = $r->nextRowWithFormula()) !== null) {
    foreach ($row as $cell) {
        $sids[] = $cell['style_id'];
        $f = $r->getStyleFormat($cell['style_id']);
        printf("sid=%d font=%s bold=%s numFmt=%s\n",
            $cell['style_id'],
            $f['font']['name'],
            var_export($f['font']['bold'], true),
            $f['format_string']);
    }
}
echo "distinct_sids: " . count(array_unique($sids)) . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/insert_text_font_with_numfmt.xlsx');
?>
--EXPECT--
sid=1 font=Arial bold=true numFmt=General
sid=2 font=Arial bold=true numFmt=#,##0.00
sid=3 font=Arial bold=true numFmt=#,##0.0000
distinct_sids: 3
