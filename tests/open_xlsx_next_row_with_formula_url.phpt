--TEST--
nextRowWithFormula adds 'url' key for hyperlink cells; SKIP_MERGED_FOLLOW nulls follow positions
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];

(new \Vtiful\Kernel\Excel($config))
    ->fileName('open_xlsx_next_row_with_formula_url.xlsx', 'S1')
    ->insertText(0, 0, 'plain')
    ->insertUrl(1, 0, 'https://example.com/x')
    ->mergeCells('A3:C3', 'M')
    ->output();

echo "rich rows (default):\n";
$r = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_next_row_with_formula_url.xlsx')
    ->openSheet();

$row = $r->nextRowWithFormula();
echo "row 0 has url? "; var_dump(array_key_exists('url', $row[0]));

$row = $r->nextRowWithFormula();
echo "row 1 url: "; var_dump($row[0]['url'] ?? null);

$row = $r->nextRowWithFormula();
echo "row 2 master only (without flag): "; var_dump(count($row));

echo "rich rows with SKIP_MERGED_FOLLOW:\n";
$r2 = (new \Vtiful\Kernel\Excel($config))
    ->openFile('open_xlsx_next_row_with_formula_url.xlsx')
    ->openSheet(null, \Vtiful\Kernel\Excel::SKIP_MERGED_FOLLOW);
$r2->nextRowWithFormula();  /* row 0 */
$r2->nextRowWithFormula();  /* row 1 */
$row = $r2->nextRowWithFormula();  /* row 2: master at A, follow at B,C — but follow only emits if cells exist in XML */
echo "row 2 cells: "; var_dump(count($row));
echo "row 2 master value: "; var_dump($row[0]['value']);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row_with_formula_url.xlsx');
?>
--EXPECT--
rich rows (default):
row 0 has url? bool(false)
row 1 url: string(21) "https://example.com/x"
row 2 master only (without flag): int(1)
rich rows with SKIP_MERGED_FOLLOW:
row 2 cells: int(1)
row 2 master value: string(1) "M"
