# Repeat columns on each printed page

Mark one or more columns as title columns that are repeated on the left of every printed page.

## Function Prototype

```php
repeatColumns(string $rangeA1): self
```

### **string $rangeA1**

> A1-style column range.
>
> A single column (e.g. `"A"`) or a closed interval (e.g. `"A:C"`).
>
> An invalid format throws an exception (error code `221`).

## Example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age', 'city'])
    ->data([
        ['viest', 21, 'Beijing'],
        ['wjx',   21, 'Shanghai']
    ])
    ->repeatColumns('A:A') // Repeat column A on every printed page
    ->output();
```
