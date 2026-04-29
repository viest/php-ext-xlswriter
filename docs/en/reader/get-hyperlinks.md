# Get hyperlinks

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns every `<hyperlink>` element in the active worksheet, covering both external URLs and intra-workbook jumps.

## Methods

```php
getHyperlinks(): ?array
```

## Return value

An array of hyperlink entries, each with:

* `first_row` *(int)* — starting row (1-based);
* `first_col` *(int)* — starting column (1-based);
* `last_row` *(int)* — ending row;
* `last_col` *(int)* — ending column;
* `url` *(string\|null)* — external URL, or `null`;
* `location` *(string\|null)* — internal target, e.g. `Sheet2!A1`;
* `display` *(string\|null)* — display text;
* `tooltip` *(string\|null)* — hover tooltip.

Returns `null` when `openSheet()` has not been called.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$links = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getHyperlinks();

print_r($links);
// Array
// (
//     [0] => Array
//         (
//             [first_row] => 1
//             [first_col] => 1
//             [last_row]  => 1
//             [last_col]  => 1
//             [url]       => https://www.php.net
//             [location]  =>
//             [display]   => PHP home
//             [tooltip]   => Click to visit
//         )
// )
```
