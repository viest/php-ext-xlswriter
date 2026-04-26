# Fit to pages

Automatically scale the worksheet so it prints inside a fixed number of pages horizontally and vertically. Common use case: squeeze a wide table down to one page wide.

## Function Prototype

```php
fitToPages(int $width, int $height): self
```

### **int $width**

> Number of pages horizontally. `1` means fit to one page wide.

### **int $height**

> Number of pages vertically. `0` means no vertical limit.

## Example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);
$fileObject = $fileObject->fileName('tutorial.xlsx');

$filePath = $fileObject->header(['name', 'age', 'city', 'email'])
    ->data([
        ['viest', 21, 'Beijing',  'viest@example.com'],
        ['wjx',   21, 'Shanghai', 'wjx@example.com']
    ])
    ->fitToPages(1, 0) // Fit to 1 page wide, unlimited height
    ->output();
```
