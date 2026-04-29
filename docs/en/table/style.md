# Style

The visual style of an Excel table is the combination of a *type* (default / light / medium / dark) and a *number* within that type. `Table::style(int $type, int $number)` takes those two values. `Table::name(string $name)` gives the table a unique name; if omitted, libxlsxwriter generates `Table1`, `Table2`, ...

## Function prototypes

```
\Vtiful\Kernel\Table::style(int $type, int $number): self
\Vtiful\Kernel\Table::name(string $name): self
```

### **int $type**

> The style family — one of:
>
> * `Table::STYLE_TYPE_DEFAULT`
> * `Table::STYLE_TYPE_LIGHT`
> * `Table::STYLE_TYPE_MEDIUM`
> * `Table::STYLE_TYPE_DARK`

### **int $number**

> The style index inside the family. Excel ships with 1..21 light, 1..28 medium and 1..11 dark variants; popular picks are `Light 11` and `Medium 9`. Pass `0` to apply no styling.

### **string $name**

> The table name. Must start with a letter or underscore and be unique inside the workbook. The name can be referenced from formulas, e.g. `Table1[@Column1]`.

## Example

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$table = new \Vtiful\Kernel\Table();
$table->name('Performance')
      ->style(\Vtiful\Kernel\Table::STYLE_TYPE_MEDIUM, 9)
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
