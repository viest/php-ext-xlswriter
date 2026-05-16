--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("18.xlsx");

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->data([['wjx', 21]]);

$filePath = $fileObject->output();

var_dump($filePath);

/* Round-trip: two consecutive ->data() calls append, not overwrite. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('18.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/18.xlsx');
?>
--EXPECT--
string(15) "./tests/18.xlsx"
array(3) {
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
  [2]=>
  array(2) {
    [0]=>
    string(3) "wjx"
    [1]=>
    int(21)
  }
}
