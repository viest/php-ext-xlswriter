# Get formula AST

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Parses a formula string into a nested abstract syntax tree (AST) so it can be inspected or rewritten in PHP.

## Methods

```php
getFormulaAst(string $formula): array
```

### **string $formula**

> Formula string. The leading `=` is optional.

## Return value

A nested tree of token nodes. Each node carries a `kind` field (`func`, `ref`, `range`, `number`, `string`, `bool`, `error`, `op`, `array`, ...) and additional fields appropriate to that kind (argument lists, reference text, operator, child nodes).

This method does not depend on workbook state; it can be called without `openFile()`.

## Example

```php
var_export(\Vtiful\Kernel\Excel::getFormulaAst('=SUM(A1:A10) + 1'));
// Output is shaped like:
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
