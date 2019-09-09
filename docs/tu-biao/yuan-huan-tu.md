# 圆环图

### 二维圆环图

![](../.gitbook/assets/chart_doughnut1.png)

```php
<?php declare(strict_types = 1);

$config = [
    'path' => './tests',
];

$dataHeader = [
    'Category', 'Values',
];

$dataRows = [
    ['Glazed', 50],
    ['Chocolate', 35],
    ['Cream', 15],
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_DOUGHNUT);

$chartResource = $chart
    // series(string $value [, string $category])
    ->series('=Sheet1!$B$2:$B$4', '=Sheet1!$A$2:$A$4')
    ->seriesName('Doughnut sales data')
    ->title('Popular Doughnut Types')
    ->style(10)
    ->toResource();

$filePath = $fileObject
    ->header($dataHeader)
    ->data($dataRows)
    ->insertChart(0, 4, $chartResource)
    ->output();
```

