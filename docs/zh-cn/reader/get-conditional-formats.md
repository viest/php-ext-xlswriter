# 条件格式

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表中所有 `<conditionalFormatting>` 区块及其规则。

## 函数原型

```php
getConditionalFormats(): ?array
```

## 返回值

返回数组，每个元素表示一个条件格式区域：

* `range` *(string)* — 应用范围，如 `A1:A10`；
* `rules` *(array)* — 规则列表，每条规则字段：
  * `type` *(string)* — 规则类型，如 `cellIs`、`expression`、`top10`、`containsText` 等；
  * `operator` *(string\|null)* — 比较运算符，如 `between`、`greaterThan`；
  * `priority` *(int)* — 优先级；
  * `stop_if_true` *(bool)* — 命中后是否停止后续规则；
  * `dxf_id` *(int\|null)* — 关联的差异格式 ID；
  * `percent` *(bool)* — `top10` 类型是否按百分比；
  * `bottom` *(bool)* — `top10` 类型是否取最小；
  * `rank` *(int\|null)* — `top10` 排名；
  * `text` *(string\|null)* — 文本类规则的匹配文本；
  * `time_period` *(string\|null)* — 时间段规则；
  * `formula1` *(string\|null)* — 公式 1；
  * `formula2` *(string\|null)* — 公式 2。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$cfs = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getConditionalFormats();

var_export($cfs);
```
