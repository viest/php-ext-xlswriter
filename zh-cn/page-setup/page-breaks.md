# 分页符

在指定行 / 列之前插入水平 / 垂直分页符，强制 Excel 在该位置开始新的一页。

## 函数原型

```php
horizontalPageBreaks(array $rows): self
verticalPageBreaks(array $cols): self
```

### **array $rows**

> 1 起始的行号数组。每个值表示在该行**之前**插入水平分页符。

### **array $cols**

> 1 起始的列号数组。每个值表示在该列**之前**插入垂直分页符。

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$rows = [];
for ($i = 1; $i <= 100; $i++) {
    $rows[] = ["row{$i}", $i];
}

$filePath = $fileObject->header(['name', 'value'])
    ->data($rows)
    ->horizontalPageBreaks([20, 40, 60, 80]) // 每 20 行一页
    ->verticalPageBreaks([3])                // 第 3 列前换页
    ->output();
```
