# Get defined names

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns every `<definedName>` element in the workbook, both workbook-level and sheet-local.

## Methods

```php
getDefinedNames(): ?array
```

## Return value

An array of name entries:

* `name` *(string)* — defined name;
* `formula` *(string)* — referenced formula or range;
* `scope` *(string\|null)* — owning worksheet name; `null` for workbook-level names;
* `hidden` *(bool)* — whether the name is hidden.

Call after `openFile()`.

## Example

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
