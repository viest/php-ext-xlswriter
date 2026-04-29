# 按页缩放

将整张工作表自动缩放到指定的横向、纵向页数范围内打印。常用于把宽表压到一页宽。

## 函数原型

```php
fitToPages(int $width, int $height): self
```

### **int $width**

> 横向页数。例如 `1` 表示宽度压缩到一页之内。

### **int $height**

> 纵向页数。例如 `0` 表示纵向不限制。

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age', 'city', 'email'])
    ->data([
        ['viest', 21, 'Beijing',  'viest@example.com'],
        ['wjx',   21, 'Shanghai', 'wjx@example.com']
    ])
    ->fitToPages(1, 0) // 宽度压到一页，纵向不限
    ->output();
```
