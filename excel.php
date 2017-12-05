<?php
$config = ['path' => './tests/'];
$excel = new \Vtiful\Kernel\Excel($config);
$fileFd = $excel->fileName('tutorial01.xlsx');
var_dump($fileFd);
$setHeader = $fileFd->header(['Item', 'Cost']);
var_dump($setHeader);
$setData = $setHeader->data([
        ['Rent', 1000],
        ['Gas',  100],
        ['Food', 300],
        ['Gym',  50],

    ]);
var_dump($setData);
$output = $setData->output();
var_dump($output);
?>

