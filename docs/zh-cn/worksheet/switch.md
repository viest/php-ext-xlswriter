# 切换工作表

## **函数原型**

```php
checkoutSheet(string $sheetName);
```

## **实例**

```php
$config = [
  'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
    ->data([
    ['viest', 21],
    ['viest', 22],
    ['viest', 23],
    ]);

// 添加工作表，并插入数据
$fileObject->addSheet('twoSheet')
    ->header(['name', 'age'])
    ->data([['vikin', 22]]);

// 切换回默认工作表，并追加数据
$fileObject->checkoutSheet('Sheet1')
    ->data([['sheet1']]);

$filePath = $fileObject->output();
```

