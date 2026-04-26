# 手动销毁资源

默认所有资源将在变量生命周期结束时全部销毁，除此之外你可以手动关闭销毁。

```php
$config = [
    'path' => '/home/viest' // xlsx文件保存路径
];

$excel = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial01.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

// 关闭当前打开的所有文件句柄 并 回收资源
$excel->close();
```

