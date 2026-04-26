# Print area

Restrict the printed output to a single rectangular range. Cells outside the range are not printed.

## Function Prototype

```php
printArea(string $rangeA1): self
```

### **string $rangeA1**

> A1-style rectangular range, e.g. `"A1:F20"`.

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
    ->printArea('A1:B3') // Only A1:B3 is printed
    ->output();
```
