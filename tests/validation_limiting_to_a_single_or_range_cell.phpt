--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
    ->valueList(['wjx', 'viest']);

$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial.xlsx')
    ->validation('A1', $validation->toResource())
    ->validation('B1:B1048576', $validation->toResource())
    ->output();

var_dump($validation, $filePath);
?>
--EXPECT--
object(Vtiful\Kernel\Validation)#1 (0) {
}
string(21) "./tests/tutorial.xlsx"

