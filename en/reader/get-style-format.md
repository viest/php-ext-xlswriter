# Get style format

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Resolves the `style_id` reported by `nextRowWithFormula()` into a complete description of the cell's formatting.

## Methods

```php
getStyleFormat(int $style_id): ?array
```

### **int $style_id**

> Style index, matching an `<xf>` entry in `xl/styles.xml`.

## Return value

The fields below are grouped by topic.

**Number format**

* `num_fmt_id` *(int)* — number format id;
* `category` *(string)* — `number`, `percent`, `date`, `time`, `datetime`, `currency`, `text`, `custom`, `general`;
* `format_string` *(string)* — format code.

**Table indices**

* `font_id` *(int)*, `fill_id` *(int)*, `border_id` *(int)*.

**Alignment / protection**

* `alignment` *(array)* — `{horizontal, vertical, wrap_text, indent, rotation}`;
* `protection` *(array)* — `{locked, hidden}`; `locked` defaults to `true`.

**Font**

* `font` *(array)* — `{name, size, color, bold, italic, strike, underline}`.

**Fill**

* `fill` *(array)* — `{pattern_type, fg_color, bg_color}`.

**Border**

* `border` *(array)* — `{left, right, top, bottom}`, each side as `{style, color}`.

Returns `null` when `style_id` is out of range.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->openFile('source.xlsx')->openSheet();

$row = $excel->nextRowWithFormula();
$sid = $row[0]['style_id'];

print_r($excel->getStyleFormat($sid));
// Array
// (
//     [num_fmt_id]    => 14
//     [category]      => date
//     [format_string] => m/d/yyyy
//     [font]          => Array ( [name] => Calibri [size] => 11 ... )
//     [fill]          => Array ( [pattern_type] => solid [fg_color] => FFFFFF00 ... )
//     [border]        => Array ( [left] => Array ( [style] => thin [color] => FF000000 ) ... )
//     [alignment]     => Array ( [horizontal] => center [vertical] => center ... )
//     [protection]    => Array ( [locked] => 1 [hidden] => 0 )
// )
```
