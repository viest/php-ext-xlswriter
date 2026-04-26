# 数据验证

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表中所有 `<dataValidation>` 元素。与 `Validation` 写入类对应。

## 函数原型

```php
getDataValidations(): ?array
```

## 返回值

返回数据验证数组，每个元素包含：

* `type` *(string)* — 验证类型，如 `whole`、`decimal`、`list`、`date`、`time`、`textLength`、`custom`；
* `operator` *(string\|null)* — 比较运算符，如 `between`、`equal`、`greaterThan`；
* `error_style` *(string\|null)* — 错误提示样式，`stop` / `warning` / `information`；
* `formula1` *(string\|null)* — 公式 1 / 列表来源；
* `formula2` *(string\|null)* — 公式 2；
* `prompt` *(string\|null)* — 输入提示内容；
* `prompt_title` *(string\|null)* — 输入提示标题；
* `error` *(string\|null)* — 错误消息内容；
* `error_title` *(string\|null)* — 错误消息标题；
* `sqref` *(string)* — 作用范围，如 `A1:A10`；
* `allow_blank` *(bool)* — 是否允许空值；
* `show_drop_down` *(bool)* — 是否显示下拉箭头（注意 XML 中此值取反）；
* `show_input_message` *(bool)* — 是否显示输入提示；
* `show_error_message` *(bool)* — 是否显示错误消息。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$dvs = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getDataValidations();

print_r($dvs);
```
