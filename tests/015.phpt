--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("15.xlsx")
    ->header(['name', 'age'])
    ->data([
        ['one', 10],
        ['two', 20],
        ['three', 30],
    ])
    ->autoFilter("A1:B3")
    ->output();

var_dump($filePath);

/* Round-trip: autoFilter doesn't affect cell values. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('15.xlsx')->openSheet()->getSheetData();
var_dump($d_);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/15.xlsx');
?>
--EXPECT--
string(15) "./tests/15.xlsx"
array(4) {
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
    string(3) "one"
    [1]=>
    int(10)
  }
  [2]=>
  array(2) {
    [0]=>
    string(3) "two"
    [1]=>
    int(20)
  }
  [3]=>
  array(2) {
    [0]=>
    string(5) "three"
    [1]=>
    int(30)
  }
}
