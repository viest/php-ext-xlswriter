# Get auto filter

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns the active worksheet's `<autoFilter>` element, or `null` if none is set.

## Methods

```php
getAutoFilter(): ?array
```

## Return value

An associative array:

* `range` *(string)* — filter range, e.g. `A1:C100`;
* `columns` *(array)* — per-column conditions; each entry has:
  * `col_id` *(int)* — column index (0-based, relative to the start of `range`);
  * `type` *(string)* — `list`, `custom`, `top10`, `dynamic`.

Type-specific extra fields:

**list**

* `values` *(string[])* — selected values.

**custom**

* `and` *(bool)* — `true` joins criteria with `AND`, `false` with `OR`;
* `criterion1` *(array)* — `{operator, value}`;
* `criterion2` *(array\|null)* — same shape, may be `null`.

**top10**

* `top` *(bool)* — `true` keeps the top N, `false` keeps the bottom N;
* `percent` *(bool)* — interpret value as a percentage;
* `value` *(float)* — N.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$af = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getAutoFilter();

var_export($af);
// array (
//   'range' => 'A1:C100',
//   'columns' => array (
//     0 => array ( 'col_id' => 0, 'type' => 'list', 'values' => array ('apple','pear') ),
//   ),
// )
```
