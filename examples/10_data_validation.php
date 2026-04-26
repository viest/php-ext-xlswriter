<?php
/* 10 — Data validation: drop-down list bound to a cell. */

$dir  = sys_get_temp_dir();
$name = '10_data_validation.xlsx';

$validation = new \Vtiful\Kernel\Validation();
$validation->validationType(\Vtiful\Kernel\Validation::TYPE_LIST)
           ->valueList(['low', 'medium', 'high']);

$excel = (new \Vtiful\Kernel\Excel(['path' => $dir]))
    ->fileName($name)
    ->header(['priority'])
    ->validation('A2', $validation->toResource());
$path = $excel->output();

echo "wrote $path — open in Excel to see the dropdown on A2\n";
@unlink($path);
