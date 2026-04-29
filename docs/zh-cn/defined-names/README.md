# 定义名称

为单元格 / 区域 / 公式起一个名字，之后即可在公式或图表里直接引用该名字，无需重复书写区域地址。

按作用域分两类：

* **工作簿级（全局）** —— 在所有工作表中可见。
* **工作表级（局部）** —— 仅在指定的工作表内可见，调用时传入 `$scopeSheet` 参数。

## 函数原型

```php
defineName(string $name, string $formula, ?string $scopeSheet = null): self
```

### **string $name**

> 名称。

### **string $formula**

> 公式字符串，**必须以 `=` 开头**，单元格 / 区域引用必须使用绝对地址（带 `$`）。
>
> 例如 `=Sheet1!$A$1` 或 `=Sheet1!$B$2:$B$10`。

### **string $scopeSheet**

> 可选。指定作用域工作表名后，名称仅在该工作表内可见；省略则为全局名称。

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx', 'Sheet1');

$filePath = $fileObject->header(['name', 'amount'])
    ->data([
        ['viest', 1000],
        ['wjx',   2000],
    ])
    ->defineName('SalesTotal', '=Sheet1!$B$2')                   // 全局名称
    ->defineName('LocalRange', '=Sheet1!$B$2:$B$3', 'Sheet1')   // 局部名称（仅 Sheet1 可见）
    ->output();
```
