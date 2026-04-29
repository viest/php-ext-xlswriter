# Get merged cells

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns every `<mergeCell>` range declared in the active worksheet.

## Methods

```php
getMergedCells(): ?array
```

## Return value

An array of merge entries, each with:

* `first_row` *(int)* — starting row (1-based);
* `first_col` *(int)* — starting column (1-based);
* `last_row` *(int)* — ending row;
* `last_col` *(int)* — ending column.

Returns an empty array when the sheet has no merges, or `null` if `openSheet()` has not been called.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$merged = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getMergedCells();

print_r($merged);
// Array
// (
//     [0] => Array
//         (
//             [first_row] => 1
//             [first_col] => 1
//             [last_row]  => 1
//             [last_col]  => 3
//         )
// )
```
