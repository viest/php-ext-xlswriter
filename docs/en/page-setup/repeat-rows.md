# Repeat rows on each printed page

Mark one or more rows as title rows that are repeated at the top of every printed page.

## Function Prototype

```php
repeatRows(string $rangeA1): self
```

### **string $rangeA1**

> A1-style row range.
>
> A single row (e.g. `"1"`) or a closed interval (e.g. `"1:3"`), 1-based.
>
> An invalid format throws an exception (error code `220`).

## Example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->repeatRows('1:1') // Repeat row 1 on every printed page
    ->output();
```
