# Worksheet list with metadata

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Unlike `sheetList()`, which only returns names, `sheetListWithMeta()` returns the visibility state of each worksheet.

## Methods

```php
sheetListWithMeta(): ?array
```

## Return value

An array with one entry per worksheet:

* `name` *(string)* — worksheet name;
* `state` *(string)* — visibility, one of `visible`, `hidden`, `veryHidden`.

Returns `null` when `openFile()` has not been called.

## Example

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
