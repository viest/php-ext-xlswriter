# Iterate charts

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Iterates every chart on a worksheet via a callback, exposing the chart type, title, anchor and data series.

## Methods

```php
iterateCharts(callable $callback, ?string $sheet = null): bool
```

### **callable $callback**

> Invoked once per chart. Return `false` from the callback to stop iteration.

### **string $sheet**

> Worksheet name. When `null`, the currently `openSheet()`-ed worksheet is used.

## Callback payload

The callback receives an associative array:

* `type` *(string)* — chart type, e.g. `bar`, `line`, `pie`, `scatter`;
* `title` *(string\|null)* — chart title;
* `anchor` *(array)* — `{from_row, from_col, to_row, to_col}` (0-based);
* `series` *(array)* — data series; each entry has:
  * `name` *(string\|null)* — series name;
  * `categories` *(string\|null)* — category reference, e.g. `Sheet1!$A$2:$A$10`;
  * `values` *(string\|null)* — values reference.

## Return value

`true` on successful completion or callback-driven early exit; `false` on error.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx');

$excel->iterateCharts(function (array $chart) {
    echo "chart {$chart['type']}, title {$chart['title']}\n";
    foreach ($chart['series'] as $s) {
        echo "  series {$s['name']}: {$s['values']}\n";
    }
}, 'Sheet1');
```
