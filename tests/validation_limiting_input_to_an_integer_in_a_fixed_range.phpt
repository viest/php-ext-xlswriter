--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_BETWEEN)
    ->minimumNumber(1)
    ->maximumNumber(10);

$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('validation_limiting_input_to_an_integer_in_a_fixed_range.xlsx')
    ->header(['Value'])
    ->validation('A1', $validation->toResource())
    ->insertText(0, 0, 20)
    ->output();

var_dump($validation, $filePath);
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/validation_limiting_input_to_an_integer_in_a_fixed_range.xlsx');
?>
--EXPECT--
object(Vtiful\Kernel\Validation)#1 (0) {
}
string(69) "./tests/validation_limiting_input_to_an_integer_in_a_fixed_range.xlsx"

