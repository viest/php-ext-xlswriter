# 定义名称

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取工作簿中所有 `<definedName>` 元素，包含工作簿级和工作表级名称。

## 函数原型

```php
getDefinedNames(): ?array
```

## 返回值

返回数组，每个元素包含：

* `name` *(string)* — 名称；
* `formula` *(string)* — 引用的公式或区域；
* `scope` *(string\|null)* — 作用域：所属工作表名；为 `null` 表示工作簿级；
* `hidden` *(bool)* — 是否隐藏。

调用前需先 `openFile()`。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$names = $excel->openFile('source.xlsx')
    ->getDefinedNames();

print_r($names);
// Array
// (
//     [0] => Array
//         (
//             [name]    => TaxRate
//             [formula] => Sheet1!$A$1
//             [scope]   =>
//             [hidden]  => false
//         )
// )
```
