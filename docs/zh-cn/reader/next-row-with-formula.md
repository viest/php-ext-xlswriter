# 公式与样式逐行读取

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

`nextRowWithFormula()` 在返回单元格值之外，附带类型、样式 ID、公式与超链接信息，是访问公式与样式的主要入口。

## 函数原型

```php
nextRowWithFormula(): ?array
```

## 返回值

迭代到末尾时返回 `null`，否则返回当前行数组，每个单元格为：

* `value` *(mixed)* — 单元格值；
* `type` *(string)* — `number`、`datetime`、`string`、`inline_string`、`boolean`、`formula`、`error`、`blank`；
* `style_id` *(int)* — 样式索引，可传给 `getStyleFormat()` 解析；
* `formula` *(string\|array\|null)* — 公式字符串；当工作簿以详尽模式打开时，该字段为数组：
  * `type` *(string)* — `normal`、`array`、`shared`、`dataTable`；
  * `text` *(string)* — 公式文本；
  * `ref` *(string\|null)* — 数组 / 共享公式区域；
  * `si` *(int\|null)* — 共享公式索引；
  * `is_dynamic` *(bool)* — 是否为动态数组公式；
  * `cached_value` *(mixed)* — XML 中缓存的计算结果；
* `url` *(string\|null)* — 单元格关联的超链接，如有。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

while (($row = $excel->nextRowWithFormula()) !== null) {
    foreach ($row as $cell) {
        echo "[{$cell['type']}] {$cell['value']} (style {$cell['style_id']})";
        if (!empty($cell['formula'])) {
            echo ' formula=' . (is_array($cell['formula']) ? $cell['formula']['text'] : $cell['formula']);
        }
        echo PHP_EOL;
    }
}
```
