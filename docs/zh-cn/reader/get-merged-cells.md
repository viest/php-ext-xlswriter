# 合并单元格

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表中所有 `<mergeCell>` 区域。

## 函数原型

```php
getMergedCells(): ?array
```

## 返回值

返回合并区域数组，每个元素包含：

* `first_row` *(int)* — 起始行号（1 起）；
* `first_col` *(int)* — 起始列号（1 起）；
* `last_row` *(int)* — 结束行号；
* `last_col` *(int)* — 结束列号。

工作表中没有任何合并单元格时返回空数组；未调用 `openSheet()` 时返回 `null`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$merged = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getMergedCells();

print_r($merged);
// Array
// (
//     [0] => Array
//         (
//             [first_row] => 1
//             [first_col] => 1
//             [last_row]  => 1
//             [last_col]  => 3
//         )
// )
```
