# 打印方向 - 纵向

## **函数原型**

```php
setPrintedPortrait(): self
```

## **实例**

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_portrait.xlsx', 'sheet1')
    ->setPrintedPortrait() // 设置打印方向为纵向
    ->output();
```