# Area map

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
    ->style(11)// Values ​​1 - 48, refer to 48 styles in the Excel 2007 Design tab
    ->axisNameX('Test number') // Set the X axis name
    ->axisNameY('Sample length (mm)') // Set the Y axis name
    ->title('Results of sample analysis') // Set the chart Title
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