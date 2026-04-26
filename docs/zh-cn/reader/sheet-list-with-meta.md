# 工作表列表（含元数据）

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

与 `sheetList()` 仅返回名称不同，`sheetListWithMeta()` 同时返回每个工作表的可见性状态。

## 函数原型

```php
sheetListWithMeta(): ?array
```

## 返回值

返回数组，每个元素描述一个工作表：

* `name` *(string)* — 工作表名称；
* `state` *(string)* — 可见性，取值为 `visible`、`hidden`、`veryHidden`。

未调用 `openFile()` 时返回 `null`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$meta = $excel->openFile('source.xlsx')
    ->sheetListWithMeta();

print_r($meta);
// Array
// (
//     [0] => Array ( [name] => Sheet1 [state] => visible )
//     [1] => Array ( [name] => Hidden  [state] => hidden  )
// )
```
