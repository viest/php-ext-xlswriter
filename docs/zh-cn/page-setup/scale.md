# 缩放比例

## 函数原型

```php
setPrintScale(int $scale): self
```

### **int $scale**

> 打印缩放比例
>
> 范围：10 <= $scale <= 400

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setPaper(\Vtiful\Kernel\Excel::PAPER_A3)
    ->setLandscape()
    ->setMargins(1, 1, 2, 2)
    ->setPrintScale(180)
    ->output();
```