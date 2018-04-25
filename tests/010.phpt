--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("excel_writer")) print "skip"; ?>
--FILE--
<?php 
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);
$fileOne = $excel->fileName('tutorial01.xlsx')
    ->header(['test1'])
    ->data([
      ['data1'],
    ])
    ->output();
$fileTwo = $excel->fileName('tutorial02.xlsx')
    ->header(['test2'])
    ->data([
        ['data2'],
    ])
    ->output();
var_dump($fileOne,$fileTwo);
?>
--EXPECT--
string(23) "./tests/tutorial01.xlsx"
string(23) "./tests/tutorial02.xlsx"