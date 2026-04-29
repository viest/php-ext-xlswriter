# 超链接

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表中所有 `<hyperlink>` 元素，包括外部 URL 与工作簿内部跳转。

## 函数原型

```php
getHyperlinks(): ?array
```

## 返回值

返回超链接数组，每个元素包含：

* `first_row` *(int)* — 起始行号（1 起）；
* `first_col` *(int)* — 起始列号（1 起）；
* `last_row` *(int)* — 结束行号；
* `last_col` *(int)* — 结束列号；
* `url` *(string\|null)* — 外部链接地址，无则为 `null`；
* `location` *(string\|null)* — 内部跳转目标，如 `Sheet2!A1`；
* `display` *(string\|null)* — 显示文本；
* `tooltip` *(string\|null)* — 鼠标悬停提示。

未调用 `openSheet()` 时返回 `null`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$links = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getHyperlinks();

print_r($links);
// Array
// (
//     [0] => Array
//         (
//             [first_row] => 1
//             [first_col] => 1
//             [last_row]  => 1
//             [last_col]  => 1
//             [url]       => https://www.php.net
//             [location]  =>
//             [display]   => PHP 官网
//             [tooltip]   => 点击访问
//         )
// )
```
