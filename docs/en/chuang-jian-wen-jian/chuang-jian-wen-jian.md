# 创建文件

```php
$config = ['path' => '/home/viest'];
$excel  = new \Vtiful\Kernel\Excel($config);

// fileName() will automatically create a worksheet,
// you can customize the worksheet name, the worksheet name is optional.
$filePath = $excel->fileName('tutorial01.xlsx', 'sheet1')
    ->header(['Item', 'Cost'])
    ->data([
        ['Rent', 1000],
        ['Gas',  100],
        ['Food', 300],
        ['Gym',  50],
    ])
    ->output();
```

