--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('show_comment.xlsx')
    ->header(['Item', 'Cost'])
    ->insertComment(0,1,'comment')
    ->showComment()
    ->output();

var_dump($filePath);

/* Round-trip: file opens and data is readable. */
$v_   = new \Vtiful\Kernel\Excel($config);
$d_   = $v_->openFile('show_comment.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/show_comment.xlsx');
?>
--EXPECT--
string(25) "./tests/show_comment.xlsx"
bool(true)
