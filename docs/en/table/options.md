# Options

These methods toggle the common feature switches on an Excel table. Each accepts an optional `bool`; passing nothing turns the option **on**, passing `false` turns it off.

## Function prototypes

```
\Vtiful\Kernel\Table::noHeaderRow(bool $on = true): self
\Vtiful\Kernel\Table::noAutofilter(bool $on = true): self
\Vtiful\Kernel\Table::noBandedRows(bool $on = true): self
\Vtiful\Kernel\Table::bandedColumns(bool $on = true): self
\Vtiful\Kernel\Table::firstColumn(bool $on = true): self
\Vtiful\Kernel\Table::lastColumn(bool $on = true): self
\Vtiful\Kernel\Table::totalRow(bool $on = true): self
```

### **noHeaderRow()**

> Hide the header row; the entire range passed to `addTable()` is treated as data.

### **noAutofilter()**

> Disable the autofilter drop-downs in the header.

### **noBandedRows()**

> Disable the alternating row banding.

### **bandedColumns()**

> Enable alternating column banding (independent of row banding).

### **firstColumn() / lastColumn()**

> Highlight the first or last column — useful for emphasising a label or summary column.

### **totalRow()**

> Add a total row as the last row of the range. Pair it with `total_string` / `total_function` on the column definitions. The range passed to `addTable()` must include the total row.

## Example

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$table = new \Vtiful\Kernel\Table();
$table->name('Sales')
      ->totalRow()
      ->bandedColumns()
      ->firstColumn()
      ->columns([
          ['header' => 'Region', 'total_string' => 'Total'],
          ['header' => 'Amount', 'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
      ]);

$excel->fileName('tutorial.xlsx')
    ->data([
        ['East', 100],
        ['West',  90],
        ['',       0], // reserved for the total row
    ])
    ->addTable('A1:B4', $table)
    ->output();
```
