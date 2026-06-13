--TEST--
Check for vtiful presence
--SKIPIF--
<?php
require __DIR__ . '/include/skipif.inc';
skip_disable_reader();
?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('11.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setColumn('A:A', 200, $boldStyle)
    ->setColumn('B:B', 200, null)
    ->output();

var_dump($filePath);

/* Round-trip: header + single data row read back exactly. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('11.xlsx')->openSheet()->getSheetData();
var_dump($d_);

/* Round-trip the column metadata: both columns share width 200, but only
 * column A has the bold style attached. The reader exposes the width via
 * getColumnOptions() and the cell-level style via nextRowWithFormula() +
 * getStyleFormat(). */
$r   = (new \Vtiful\Kernel\Excel($config))->openFile('11.xlsx')->openSheet();
$colA = $r->getColumnOptions('A');
$colB = $r->getColumnOptions('B');
echo "colA.width: " . $colA['width'] . "\n";
echo "colB.width: " . $colB['width'] . "\n";
echo "widths_equal: " . var_export($colA['width'] === $colB['width'], true) . "\n";

$row0 = $r->nextRowWithFormula();
$row1 = $r->nextRowWithFormula();
echo "A.styleId: " . $row0[0]['style_id'] . "\n";
echo "B.styleId: " . $row0[1]['style_id'] . "\n";
echo "A.bold: " . var_export($r->getStyleFormat($row0[0]['style_id'])['font']['bold'], true) . "\n";
echo "rows_styles_consistent: " . var_export(
    $row0[0]['style_id'] === $row1[0]['style_id'] && $row0[1]['style_id'] === $row1[1]['style_id'], true
) . "\n";
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/11.xlsx');
?>
--EXPECT--
string(15) "./tests/11.xlsx"
array(2) {
  [0]=>
  array(2) {
    [0]=>
    string(4) "name"
    [1]=>
    string(3) "age"
  }
  [1]=>
  array(2) {
    [0]=>
    string(5) "viest"
    [1]=>
    int(21)
  }
}
colA.width: 200.7109375
colB.width: 200.7109375
widths_equal: true
A.styleId: 1
B.styleId: 0
A.bold: true
rows_styles_consistent: true
