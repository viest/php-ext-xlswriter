--TEST--
Check for vtiful presence
--SKIPIF--
<?php if (!extension_loaded("xlswriter")) print "skip"; ?>
--FILE--
<?php
$config = ['path' => './tests'];

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_INTEGER)
    ->criteriaType(\Vtiful\Kernel\Validation::CRITERIA_GREATER_THAN)
    ->valueNumber(20);

$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('validation_limiting_input_to_an_integer_greater_than_a_fixed_value.xlsx')
    ->validation('A1', $validation->toResource())
    ->insertText(0, 0, 21)
    ->output();

var_dump($validation, $filePath);

/* Round-trip: validation didn't corrupt the workbook. */
$v_ = new \Vtiful\Kernel\Excel($config);
$d_ = $v_->openFile('validation_limiting_input_to_an_integer_greater_than_a_fixed_value.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/validation_limiting_input_to_an_integer_greater_than_a_fixed_value.xlsx');
?>
--EXPECTF--
object(Vtiful\Kernel\Validation)#%d (0) {
}
string(79) "./tests/validation_limiting_input_to_an_integer_greater_than_a_fixed_value.xlsx"
bool(true)
