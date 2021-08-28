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

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('open_xlsx_next_row_with_set_type.xlsx');

$filePath = $fileObject->data([
        [1, 'Test']
    ])
    ->insertDate(1, 2, 1568389354, 'mmm d yyyy hh:mm AM/PM')
    ->output();

$fileObject->openFile('open_xlsx_next_row_with_set_type.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_INT,
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
    ]);

var_dump($fileObject->nextRow()); // Header
var_dump($fileObject->nextRow());
var_dump($fileObject->nextRow());
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/open_xlsx_next_row_with_set_type.xlsx');
?>
--EXPECT--
array(2) {
  [0]=>
  int(1)
  [1]=>
  string(4) "Test"
}
array(3) {
  [0]=>
  NULL
  [1]=>
  string(0) ""
  [2]=>
  int(1568389354)
}
NULL