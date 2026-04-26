# Columns

`Table::columns(array $columns)` configures the per-column metadata. `$columns` is an array; each element is an associative array whose keys describe one column. Columns left out fall back to libxlsxwriter's defaults (`Column1`, `Column2`, ...).

## Function prototype

```
\Vtiful\Kernel\Table::columns(array $columns): self
```

### **array $columns**

> Each entry supports the following keys:
>
> * **header** *(string)* — column title, replaces the default `ColumnN`.
> * **formula** *(string)* — column formula applied to every data row, e.g. `'=SUM(Table1[@[Q1]:[Q4]])'`.
> * **total_string** *(string)* — string shown in the total row for this column, e.g. `'Total'`.
> * **total_function** *(int)* — total-row aggregator, one of `Table::FUNCTION_*`.
> * **total_value** *(float)* — explicit numeric value for the total row (normally derived from `total_function`).
> * **format** *(Format|resource)* — cell format for the data cells of this column. Accepts a `Format` instance or its resource handle.
> * **header_format** *(Format|resource)* — format for the header cell.

`Table::FUNCTION_*` constants: `FUNCTION_NONE`, `FUNCTION_AVERAGE`, `FUNCTION_COUNT_NUMS`, `FUNCTION_COUNT`, `FUNCTION_MAX`, `FUNCTION_MIN`, `FUNCTION_STD_DEV`, `FUNCTION_SUM`, `FUNCTION_VAR`.

## Example

A sales table with four quarterly columns and a total row.

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileHandle = $excel->fileName('tutorial.xlsx')->getHandle();
$bold       = new \Vtiful\Kernel\Format($fileHandle);
$boldFmt    = $bold->bold()->toResource();

$table = new \Vtiful\Kernel\Table();
$table->name('Sales')
      ->totalRow()
      ->columns([
          ['header' => 'Region', 'header_format' => $boldFmt, 'total_string' => 'Total'],
          ['header' => 'Q1',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
          ['header' => 'Q2',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
          ['header' => 'Q3',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
          ['header' => 'Q4',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
      ]);

$excel->data([
        ['East',  100, 110, 120, 130],
        ['West',   90,  95, 100, 105],
        ['North', 200, 210, 220, 230],
        // last row is reserved for the total row
        ['',        0,   0,   0,   0],
    ])
    ->addTable('A1:E5', $table)
    ->output();
```
