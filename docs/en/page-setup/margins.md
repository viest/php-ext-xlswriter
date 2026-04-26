# Margins

## Function Prototype

```php
setMargins(double $left, double $right, double $top, double $bottom): self
```

The unit is inches.

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
    ->setPaper(\Vtiful\Kernel\Excel::PAPER_A3)
    ->setLandscape()
    ->setMargins(1, 1, 2, 2)
    ->output();
```
