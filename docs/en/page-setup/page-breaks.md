# Page breaks

Insert horizontal / vertical page breaks before specific rows / columns to force Excel to start a new page at that position.

## Function Prototype

```php
horizontalPageBreaks(array $rows): self
verticalPageBreaks(array $cols): self
```

### **array $rows**

> 1-based row numbers. Each value inserts a horizontal break **before** that row.

### **array $cols**

> 1-based column numbers. Each value inserts a vertical break **before** that column.

## Example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$rows = [];
for ($i = 1; $i <= 100; $i++) {
    $rows[] = ["row{$i}", $i];
}

$filePath = $fileObject->header(['name', 'value'])
    ->data($rows)
    ->horizontalPageBreaks([20, 40, 60, 80]) // New page every 20 rows
    ->verticalPageBreaks([3])                // Page break before column 3
    ->output();
```
