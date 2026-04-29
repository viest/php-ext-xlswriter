# Get sheet protection

* Extension version `1.6.0` or later;
* Build with `--enable-reader`;

Returns the active worksheet's `<sheetProtection>` element. Returns `null` when no sheet protection is enabled.

## Methods

```php
getSheetProtection(): ?array
```

## Return value

An associative array, grouped below by topic.

**Password**

* `password_hash` *(string\|null)* — hash from `<sheetProtection password="...">`.

**Protection scope**

* `sheet` *(bool)* — protect the sheet itself;
* `content` *(bool)* — protect sheet content;
* `objects` *(bool)* — protect drawing objects;
* `scenarios` *(bool)* — protect scenarios.

**Editing permissions**

* `format_cells` *(bool)* — allow cell formatting;
* `format_columns` *(bool)* — allow column formatting;
* `format_rows` *(bool)* — allow row formatting;
* `insert_columns` *(bool)* — allow inserting columns;
* `insert_rows` *(bool)* — allow inserting rows;
* `insert_hyperlinks` *(bool)* — allow inserting hyperlinks;
* `delete_columns` *(bool)* — allow deleting columns;
* `delete_rows` *(bool)* — allow deleting rows.

**Selection / other**

* `select_locked_cells` *(bool)* — allow selecting locked cells;
* `select_unlocked_cells` *(bool)* — allow selecting unlocked cells;
* `sort` *(bool)* — allow sorting;
* `auto_filter` *(bool)* — allow auto-filter;
* `pivot_tables` *(bool)* — allow pivot tables.

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$prot = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getSheetProtection();

var_export($prot);
```
