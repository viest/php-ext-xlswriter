# 自动筛选

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表的 `<autoFilter>` 元素。无自动筛选时返回 `null`。

## 函数原型

```php
getAutoFilter(): ?array
```

## 返回值

返回数组，包含：

* `range` *(string)* — 筛选范围，如 `A1:C100`；
* `columns` *(array)* — 各列筛选条件，每个元素包含：
  * `col_id` *(int)* — 列序号（0 起，相对 `range` 起始列）；
  * `type` *(string)* — `list`、`custom`、`top10`、`dynamic`。

不同 `type` 携带不同字段：

**list**

* `values` *(string[])* — 选中的值列表。

**custom**

* `and` *(bool)* — 两个条件之间的连接符，`true` 为 `AND`，`false` 为 `OR`；
* `criterion1` *(array)* — `{operator, value}`；
* `criterion2` *(array\|null)* — 同上，可能为 `null`。

**top10**

* `top` *(bool)* — `true` 取前 N，`false` 取后 N；
* `percent` *(bool)* — 是否按百分比；
* `value` *(float)* — 数值。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$af = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getAutoFilter();

var_export($af);
// array (
//   'range' => 'A1:C100',
//   'columns' => array (
//     0 => array ( 'col_id' => 0, 'type' => 'list', 'values' => array ('apple','pear') ),
//   ),
// )
```
