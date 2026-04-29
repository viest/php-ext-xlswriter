# Row / column options

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Read sizing, hidden state and outline metadata for a single row / column, plus the worksheet defaults.

## Methods

```php
getRowOptions(int $row): ?array
getColumnOptions(string $column): ?array
getDefaultRowHeight(): ?float
getDefaultColumnWidth(): ?float
```

### **int $row**

> Row index, 0-based.

### **string $column**

> Column letter, e.g. `"A"`, `"AB"`.

## Return value

`getRowOptions()` returns:

* `height` *(float\|null)* — row height in points; `null` when not customised;
* `hidden` *(bool)* — whether the row is hidden;
* `outline_level` *(int)* — outline level;
* `collapsed` *(bool)* — whether the outline is collapsed;
* `custom_height` *(bool)* — whether a custom height was set.

`getColumnOptions()` returns:

* `width` *(float\|null)* — column width in characters;
* `hidden` *(bool)* — whether the column is hidden;
* `outline_level` *(int)* — outline level;
* `collapsed` *(bool)* — whether the outline is collapsed.

`getDefaultRowHeight()` / `getDefaultColumnWidth()` return the values declared in `<sheetFormatPr>`, or `null` if none are present.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

print_r($excel->getRowOptions(0));
// Array ( [height] => 24 [hidden] => false [outline_level] => 0 [collapsed] => false [custom_height] => true )

print_r($excel->getColumnOptions('B'));
// Array ( [width] => 18 [hidden] => false [outline_level] => 0 [collapsed] => false )

var_dump($excel->getDefaultRowHeight());   // float(15)
var_dump($excel->getDefaultColumnWidth()); // float(8.43) or NULL
```
