--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
    ->valueList(['wjx', 'viest'])
    ->valueList(['wjx', 'viest']);

$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('validation_limiting_input_to_a_value_in_a_dropdown_list.xlsx')
    ->validation('A1', $validation->toResource())
    ->output();

var_dump($validation, $filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/validation_limiting_input_to_a_value_in_a_dropdown_list.xlsx');
?>
--EXPECT--
object(Vtiful\Kernel\Validation)#1 (0) {
}
string(68) "./tests/validation_limiting_input_to_a_value_in_a_dropdown_list.xlsx"

