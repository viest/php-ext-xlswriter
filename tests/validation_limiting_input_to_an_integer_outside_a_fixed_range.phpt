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
    ->minimumFormula('=A1')
    ->maximumFormula('=B1');

$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial.xlsx')
    ->header([1, 10])
    ->validation('C1', $validation->toResource())
    ->insertText(0, 2, 20)
    ->output();

var_dump($validation, $filePath);
?>
--EXPECT--
object(Vtiful\Kernel\Validation)#1 (0) {
}
string(21) "./tests/tutorial.xlsx"

