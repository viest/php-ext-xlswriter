# 遍历图片

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

通过回调遍历工作表中的所有内嵌图片，回调中可直接拿到二进制内容。

## 函数原型

```php
iterateImages(callable $callback, ?string $sheet = null): bool
```

### **callable $callback**

> 每张图片调用一次。回调返回 `false` 时中断遍历。

### **string $sheet**

> 工作表名称；为 `null` 时使用当前已 `openSheet()` 的工作表。

## 回调参数

回调接收一个数组，字段：

* `from_row`、`from_col` *(int)* — 锚点起始单元格（0 起）；
* `to_row`、`to_col` *(int)* — 锚点结束单元格；
* `mime` *(string)* — MIME 类型，如 `image/png`、`image/jpeg`；
* `data` *(string)* — 二进制数据；
* `name` *(string)* — `xl/media/` 中的文件名。

## 返回值

成功完成或被回调中断返回 `true`；工作表打开失败等错误返回 `false`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx');

$excel->iterateImages(function (array $img) {
    echo "图片位于 ({$img['from_row']}, {$img['from_col']}), MIME={$img['mime']}\n";
    file_put_contents('./out_' . $img['name'], $img['data']);
}, 'Sheet1');
```
