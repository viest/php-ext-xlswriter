# 重复打印列

将指定的列设置为标题列，每张打印页面的左侧都会重复打印这些列。

## 函数原型

```php
repeatColumns(string $rangeA1): self
```

### **string $rangeA1**

> A1 表示法的列范围字符串。
>
> 可以是单列（如 `"A"`）或闭区间（如 `"A:C"`）。
>
> 传入非法格式时抛出异常（错误码 `221`）。

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age', 'city'])
    ->data([
        ['viest', 21, 'Beijing'],
        ['wjx',   21, 'Shanghai']
    ])
    ->repeatColumns('A:A') // A 列作为标题列，每页左侧重复
    ->output();
```
