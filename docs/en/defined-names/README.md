# Defined names

Assign a name to a cell, range, or formula so it can be referenced by name from other formulas or charts instead of by address.

Two scopes are supported:

* **Workbook (global)** — visible from every worksheet.
* **Worksheet (local)** — visible only inside the worksheet whose name is passed as `$scopeSheet`.

## Function Prototype

```php
defineName(string $name, string $formula, ?string $scopeSheet = null): self
```

### **string $name**

> The name to define.

### **string $formula**

> Formula string. **Must start with `=`** and cell / range references must be absolute (use `$`).
>
> e.g. `=Sheet1!$A$1` or `=Sheet1!$B$2:$B$10`.

### **string $scopeSheet**

> Optional. When given, the name is local to that worksheet; otherwise the name is global.

## Example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx', 'Sheet1');

$filePath = $fileObject->header(['name', 'amount'])
    ->data([
        ['viest', 1000],
        ['wjx',   2000],
    ])
    ->defineName('SalesTotal', '=Sheet1!$B$2')                   // global name
    ->defineName('LocalRange', '=Sheet1!$B$2:$B$3', 'Sheet1')    // local to Sheet1
    ->output();
```
