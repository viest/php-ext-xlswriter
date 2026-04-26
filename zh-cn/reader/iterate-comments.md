# 遍历批注

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

通过回调遍历工作表中的所有批注，包括传统批注与线程化（threaded）批注。

## 函数原型

```php
iterateComments(callable $callback, ?string $sheet = null): bool
```

### **callable $callback**

> 每条批注调用一次。回调返回 `false` 时中断遍历。

### **string $sheet**

> 工作表名称；为 `null` 时使用当前已 `openSheet()` 的工作表。

## 回调参数

回调接收一个数组，字段：

* `row` *(int)* — 行号（0 起）；
* `col` *(int)* — 列号（0 起）；
* `text` *(string)* — 批注文本；
* `author` *(string\|null)* — 作者；
* `visible` *(bool)* — 是否始终显示；
* `threaded` *(bool)* — 是否为线程化批注；
* `parent_id` *(string\|null)* — 线程化批注的父级 ID。

## 返回值

成功完成或被回调中断返回 `true`，否则 `false`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx');

$excel->iterateComments(function (array $c) {
    echo "[{$c['author']}] R{$c['row']}C{$c['col']}: {$c['text']}\n";
}, 'Sheet1');
```
