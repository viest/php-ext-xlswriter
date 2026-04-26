# 打印区域

指定一个矩形区域作为打印输出范围，区域之外的单元格不会出现在打印结果中。

## 函数原型

```php
printArea(string $rangeA1): self
```

### **string $rangeA1**

> A1 表示法的矩形区域字符串，例如 `"A1:F20"`。

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
    ->printArea('A1:B3') // 仅打印 A1:B3 区域
    ->output();
```
