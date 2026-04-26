# 行列尺寸与可见性

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取指定行或列的尺寸、隐藏、分级显示等属性，以及工作表默认行高与列宽。

## 函数原型

```php
getRowOptions(int $row): ?array
getColumnOptions(string $column): ?array
getDefaultRowHeight(): ?float
getDefaultColumnWidth(): ?float
```

### **int $row**

> 行号，0 起。

### **string $column**

> 列字母，如 `"A"`、`"AB"`。

## 返回值

`getRowOptions()` 返回数组：

* `height` *(float\|null)* — 行高（点）；未自定义时为 `null`；
* `hidden` *(bool)* — 是否隐藏；
* `outline_level` *(int)* — 分级显示级别；
* `collapsed` *(bool)* — 是否折叠；
* `custom_height` *(bool)* — 是否使用自定义行高。

`getColumnOptions()` 返回数组：

* `width` *(float\|null)* — 列宽（字符宽）；
* `hidden` *(bool)* — 是否隐藏；
* `outline_level` *(int)* — 分级显示级别；
* `collapsed` *(bool)* — 是否折叠。

`getDefaultRowHeight()` / `getDefaultColumnWidth()` 返回工作表 `<sheetFormatPr>` 中声明的默认值，未声明时返回 `null`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

print_r($excel->getRowOptions(0));
// Array ( [height] => 24 [hidden] => false [outline_level] => 0 [collapsed] => false [custom_height] => true )

print_r($excel->getColumnOptions('B'));
// Array ( [width] => 18 [hidden] => false [outline_level] => 0 [collapsed] => false )

var_dump($excel->getDefaultRowHeight());   // float(15)
var_dump($excel->getDefaultColumnWidth()); // float(8.43) 或 NULL
```
