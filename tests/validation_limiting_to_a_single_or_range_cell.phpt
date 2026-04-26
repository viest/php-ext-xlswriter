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

/* Round-trip: validation didn't corrupt the workbook. */
$v_ = new \Vtiful\Kernel\Excel($config);
$d_ = $v_->openFile('tutorial.xlsx')->openSheet()->getSheetData();
var_dump(is_array($d_));
?>
--CLEAN--
<?php
@unlink(__DIR__ . '/tutorial.xlsx');
?>
--EXPECTF--
object(Vtiful\Kernel\Validation)#%d (0) {
}
string(21) "./tests/tutorial.xlsx"
bool(true)
