# 富文本逐行读取

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

`nextRowRich()` 与 `nextRow()` 类似，但字符串单元格会展开为富文本片段数组，保留每段文字的字体属性。

## 函数原型

```php
nextRowRich(): ?array
```

## 返回值

返回当前行的单元格数组：迭代到末尾时返回 `null`。

* 字符串 / 共享字符串单元格 → 数组，每个元素为一个文本片段：
  * `text` *(string)* — 片段文本；
  * `font` *(array)* — `{name, size, bold, italic, strike, underline, color}`，未声明字段为 `null`。
* 其他类型（数字、日期、布尔、错误、空）→ 与 `nextRow()` 相同的原始值。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

while (($row = $excel->nextRowRich()) !== null) {
    print_r($row);
}
// 一行示例：
// Array
// (
//     [0] => Array
//         (
//             [0] => Array
//                 (
//                     [text] => Hello
//                     [font] => Array ( [name] => Calibri [size] => 11 [bold] => 1 [italic] => 0 ... )
//                 )
//             [1] => Array ( [text] =>  world, [font] => Array (...) )
//         )
//     [1] => 42        // 数字单元格保持原值
// )
```
