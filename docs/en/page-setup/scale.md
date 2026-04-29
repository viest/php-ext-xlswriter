# Print scale

## Function Prototype

```php
setPrintScale(int $scale): self
```

### **int $scale**

> Printing scale percentage.
>
> Range: 10 <= $scale <= 400

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
    ->setPrintScale(180)
    ->output();
```
