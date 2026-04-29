# Excel table

An Excel table marks a rectangular region of a worksheet as a "smart table" with header row, autofilter, banded rows, total row and so on. The cells themselves stay normal cells; the table just adds a metadata file under `xl/tables/tableN.xml`.

There are two steps:

1. Build the options with `Vtiful\Kernel\Table`.
2. Apply them with `Excel::addTable(string $rangeA1, ?Table $opts = null)`.

> Note: `Table` does **not** use `toResource()`. Pass the `Table` instance directly as the second argument to `addTable()`. Omit the second argument to use libxlsxwriter's default table style.

## Function prototype

```
\Vtiful\Kernel\Excel::addTable(string $rangeA1, ?\Vtiful\Kernel\Table $opts = null): self
```

### **string $rangeA1**

> The cell range covered by the table, in A1 notation, e.g. `A1:D11`. The range must include both the header row and at least one data row.

### **?Table $opts**

> Optional `Table` builder. When `null`, libxlsxwriter's defaults are used.

## Sub-sections

* [Columns](columns.md)
* [Style](style.md)
* [Options](options.md)

## Example

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$table = new \Vtiful\Kernel\Table();
$table->name('Performance')
      ->style(\Vtiful\Kernel\Table::STYLE_TYPE_LIGHT, 11)
      ->columns([
          ['header' => 'Name'],
          ['header' => 'Score'],
      ]);

$excel->fileName('tutorial.xlsx')
    ->data([
        ['Alice', 90],
        ['Bob',   80],
    ])
    ->addTable('A1:B3', $table)
    ->output();
```
