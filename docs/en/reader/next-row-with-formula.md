# Next row with formula

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

`nextRowWithFormula()` returns the current row with cell values, types, style ids, formula details and hyperlinks. It is the primary entry point for inspecting formulas and styles.

## Methods

```php
nextRowWithFormula(): ?array
```

## Return value

`null` when the iterator is exhausted; otherwise an array of cells, each shaped as:

* `value` *(mixed)* — cell value;
* `type` *(string)* — `number`, `datetime`, `string`, `inline_string`, `boolean`, `formula`, `error`, `blank`;
* `style_id` *(int)* — style index, can be passed to `getStyleFormat()`;
* `formula` *(string\|array\|null)* — formula text. When the workbook is opened in verbose mode this becomes an array:
  * `type` *(string)* — `normal`, `array`, `shared`, `dataTable`;
  * `text` *(string)* — formula text;
  * `ref` *(string\|null)* — array / shared formula range;
  * `si` *(int\|null)* — shared formula index;
  * `is_dynamic` *(bool)* — dynamic-array formula;
  * `cached_value` *(mixed)* — value cached in XML;
* `url` *(string\|null)* — hyperlink attached to the cell, if any.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

while (($row = $excel->nextRowWithFormula()) !== null) {
    foreach ($row as $cell) {
        echo "[{$cell['type']}] {$cell['value']} (style {$cell['style_id']})";
        if (!empty($cell['formula'])) {
            echo ' formula=' . (is_array($cell['formula']) ? $cell['formula']['text'] : $cell['formula']);
        }
        echo PHP_EOL;
    }
}
```
