# 创建文件

> 如果路径下有相同命名的文件，新文件会覆盖老文件

```php
$config = [
    'path' => '/home/viest' // xlsx文件保存路径
];
$excel  = new \Vtiful\Kernel\Excel($config);

// fileName 会自动创建一个工作表，你可以自定义该工作表名称，工作表名称为可选参数
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

