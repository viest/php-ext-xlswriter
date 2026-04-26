# Next row with rich text

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

`nextRowRich()` behaves like `nextRow()`, but string cells are returned as an array of rich-text runs that preserve per-segment font attributes.

## Methods

```php
nextRowRich(): ?array
```

## Return value

The current row's cells, or `null` when the iterator is exhausted.

* String / shared-string cells become an array of runs:
  * `text` *(string)* — run text;
  * `font` *(array)* — `{name, size, bold, italic, strike, underline, color}`; missing attributes are `null`.
* All other cell types (number, datetime, boolean, error, blank) keep the same raw value as `nextRow()`.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

while (($row = $excel->nextRowRich()) !== null) {
    print_r($row);
}
// Example row:
// Array
// (
//     [0] => Array
//         (
//             [0] => Array
//                 (
//                     [text] => Hello
//                     [font] => Array ( [name] => Calibri [size] => 11 [bold] => 1 [italic] => 0 ... )
//                 )
//             [1] => Array ( [text] =>  world, [font] => Array (...) )
//         )
//     [1] => 42        // numeric cells keep their raw value
// )
```
