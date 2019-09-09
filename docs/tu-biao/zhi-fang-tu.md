# 直方图

![](../.gitbook/assets/chart_simple.png)

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_COLUMN);

$chartResource = $chart->series('Sheet1!$A$1:$A$5')
    ->series('Sheet1!$B$1:$B$5')
    ->series('Sheet1!$C$1:$C$5')
    ->toResource();

$filePath = $fileObject->data([
    [1, 2, 3],
    [2, 4, 6],
    [3, 6, 9],
    [4, 8, 12],
    [5, 10, 15],
])->insertChart(0, 3, $chartResource)->output();
```

