# 样式格式解析

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

将单元格的 `style_id`（来自 `nextRowWithFormula()`）解析为完整的格式描述。

## 函数原型

```php
getStyleFormat(int $style_id): ?array
```

### **int $style_id**

> 单元格样式索引，对应 `xl/styles.xml` 中的 `<xf>` 项。

## 返回值

返回字段按主题分组：

**编号格式**

* `num_fmt_id` *(int)* — 数字格式 ID；
* `category` *(string)* — `number`、`percent`、`date`、`time`、`datetime`、`currency`、`text`、`custom`、`general`；
* `format_string` *(string)* — 格式串。

**关联表索引**

* `font_id` *(int)*、`fill_id` *(int)*、`border_id` *(int)*。

**对齐 / 保护**

* `alignment` *(array)* — `{horizontal, vertical, wrap_text, indent, rotation}`；
* `protection` *(array)* — `{locked, hidden}`，默认 `locked = true`。

**字体**

* `font` *(array)* — `{name, size, color, bold, italic, strike, underline}`。

**填充**

* `fill` *(array)* — `{pattern_type, fg_color, bg_color}`。

**边框**

* `border` *(array)* — `{left, right, top, bottom}`，每个方向为 `{style, color}`。

`style_id` 越界时返回 `null`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

$row = $excel->nextRowWithFormula();
$sid = $row[0]['style_id'];

print_r($excel->getStyleFormat($sid));
// Array
// (
//     [num_fmt_id]    => 14
//     [category]      => date
//     [format_string] => m/d/yyyy
//     [font]          => Array ( [name] => Calibri [size] => 11 ... )
//     [fill]          => Array ( [pattern_type] => solid [fg_color] => FFFFFF00 ... )
//     [border]        => Array ( [left] => Array ( [style] => thin [color] => FF000000 ) ... )
//     [alignment]     => Array ( [horizontal] => center [vertical] => center ... )
//     [protection]    => Array ( [locked] => 1 [hidden] => 0 )
// )
```
