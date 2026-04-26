# 面积图

![](../../.gitbook/assets/chart_area1.png)

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$chart = new \Vtiful\Kernel\Chart($fileHandle, \Vtiful\Kernel\Chart::CHART_AREA);

$chartResource = $chart
    ->series('=Sheet1!$B$2:$B$7', '=Sheet1!$A$2:$A$7')
    ->seriesName('=Sheet1!$B$1')
    ->series('=Sheet1!$C$2:$C$7', '=Sheet1!$A$2:$A$7')
    ->seriesName('=Sheet1!$C$1')
    ->style(11)// 值为 1 - 48，可参考 Excel 2007 "设计" 选项卡中的 48 种样式
    ->axisNameX('Test number') // 设置 X 轴名称
    ->axisNameY('Sample length (mm)') // 设置 Y 轴名称
    ->title('Results of sample analysis') // 设置图表 Title
    ->toResource();

$filePath = $fileObject->header(['Number', 'Batch 1', 'Batch 2'])
    ->data([
        [2, 40, 30],
        [3, 40, 25],
        [4, 50, 30],
        [5, 30, 10],
        [6, 25, 5],
        [7, 50, 10],
    ])->insertChart(0, 3, $chartResource)->output();
```

