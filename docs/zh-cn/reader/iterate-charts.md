# 遍历图表

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

通过回调遍历工作表中的所有图表，提取类型、标题、锚点位置以及各数据系列。

## 函数原型

```php
iterateCharts(callable $callback, ?string $sheet = null): bool
```

### **callable $callback**

> 每个图表调用一次。回调返回 `false` 时中断遍历。

### **string $sheet**

> 工作表名称；为 `null` 时使用当前已 `openSheet()` 的工作表。

## 回调参数

回调接收一个数组，字段：

* `type` *(string)* — 图表类型，如 `bar`、`line`、`pie`、`scatter`；
* `title` *(string\|null)* — 图表标题；
* `anchor` *(array)* — `{from_row, from_col, to_row, to_col}`，0 起；
* `series` *(array)* — 数据系列：
  * `name` *(string\|null)* — 系列名称；
  * `categories` *(string\|null)* — 类别引用，如 `Sheet1!$A$2:$A$10`；
  * `values` *(string\|null)* — 数据引用。

## 返回值

成功完成或被回调中断返回 `true`，否则 `false`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx');

$excel->iterateCharts(function (array $chart) {
    echo "图表类型 {$chart['type']}, 标题 {$chart['title']}\n";
    foreach ($chart['series'] as $s) {
        echo "  系列 {$s['name']}: {$s['values']}\n";
    }
}, 'Sheet1');
```
