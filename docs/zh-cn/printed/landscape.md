# 打印方向 - 横向

## **函数原型**

```php
setPrintedLandscape(): self
```

## **实例**

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_landscape.xlsx', 'sheet1')
    ->setPrintedLandscape() // 设置打印方向为横向
    ->output();
```