# 重复打印行

将指定的行设置为标题行，每张打印页面的顶部都会重复打印这些行。

## 函数原型

```php
repeatRows(string $rangeA1): self
```

### **string $rangeA1**

> A1 表示法的行范围字符串。
>
> 可以是单行（如 `"1"`）或闭区间（如 `"1:3"`），均为 1 起始。
>
> 传入非法格式时抛出异常（错误码 `220`）。

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
    ->repeatRows('1:1') // 第 1 行作为标题行，每页顶部重复
    ->output();
```
