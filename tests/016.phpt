--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("16.xlsx")
    ->mergeCells('A1:C1', 'Merge cells')
    ->mergeCells('A2:C2', 'Merge cells, explicit null (default) format', null)
    ->output();

var_dump($filePath);

/* Round-trip: merged cells expose the anchor value (top-left), the rest
   of the merged range stays empty in the underlying XML. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('16.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/16.xlsx');
?>
--EXPECT--
string(15) "./tests/16.xlsx"
array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(11) "Merge cells"
  }
  [1]=>
  array(1) {
    [0]=>
    string(43) "Merge cells, explicit null (default) format"
  }
}
