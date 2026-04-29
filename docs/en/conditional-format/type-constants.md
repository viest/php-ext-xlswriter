# Type constants

Constants accepted by `type()` to declare the overall kind of conditional format rule.

## Defining class

```php
\Vtiful\Kernel\ConditionalFormat
```

## Constants

```php
const TYPE_CELL          = ?; // Cell-value rule (combine with criteria + value/minimum/maximum)
const TYPE_TEXT          = ?; // Text rule (contains / not contains / begins with / ends with)
const TYPE_TIME_PERIOD   = ?; // Date/time rule (yesterday / today / this week / etc.)
const TYPE_AVERAGE       = ?; // Average rule (above / below average)
const TYPE_DUPLICATE     = ?; // Duplicate values
const TYPE_UNIQUE        = ?; // Unique values
const TYPE_TOP           = ?; // Top N or top N%
const TYPE_BOTTOM        = ?; // Bottom N or bottom N%
const TYPE_BLANKS        = ?; // Blank cells
const TYPE_NO_BLANKS     = ?; // Non-blank cells
const TYPE_ERRORS        = ?; // Error values
const TYPE_NO_ERRORS     = ?; // Non-error values
const TYPE_FORMULA       = ?; // Custom formula (set with valueString)
const TYPE_2_COLOR_SCALE = ?; // Two-colour scale
const TYPE_3_COLOR_SCALE = ?; // Three-colour scale
const TYPE_DATA_BAR      = ?; // Data bar
const TYPE_ICON_SETS     = ?; // Icon set
```

> The numeric values are decided by the extension internally. Always reference these constants by name; do not hard-code the integers.

## Example

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
   ->value(50);
```
