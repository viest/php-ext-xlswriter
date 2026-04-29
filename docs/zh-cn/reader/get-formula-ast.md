# 公式 AST

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

将公式字符串解析为嵌套的抽象语法树（AST），便于在 PHP 侧分析或改写公式。

## 函数原型

```php
getFormulaAst(string $formula): array
```

### **string $formula**

> 公式字符串，开头的 `=` 可有可无。

## 返回值

返回嵌套的 token 节点数组，每个节点包含一个 `kind` 字段（如 `func`、`ref`、`range`、`number`、`string`、`bool`、`error`、`op`、`array`），并按节点类型携带子字段（参数列表、引用文本、运算符、子节点等）。

该方法不依赖工作簿状态，无需调用 `openFile()` 即可使用。

## 示例

```php
var_export(\Vtiful\Kernel\Excel::getFormulaAst('=SUM(A1:A10) + 1'));
// 输出形如：
// array (
//   'kind' => 'op',
//   'op'   => '+',
//   'args' => array (
//     0 => array (
//       'kind' => 'func',
//       'name' => 'SUM',
//       'args' => array ( 0 => array ('kind' => 'range', 'text' => 'A1:A10') ),
//     ),
//     1 => array ('kind' => 'number', 'value' => 1.0),
//   ),
// )
```
