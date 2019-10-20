# 缩放比例

## **函数原型**

```php
zoom(int $scale = 100): self
```

### **int $scale**

> 工作表缩放
>
> 范围：10 <= $scale <= 400
>
> 默认值： 100
>
> 缩放比例不影响打印比例



## **实例**

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
    ->zoom(200) // 设置工作表缩放系数
    ->data([
    ['viest', 21],
    ['viest', 22],
    ['viest', 23],
    ])
    ->output();
```

