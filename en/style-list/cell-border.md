# Cell border

### **Function Prototype**

```php
Border(int $borderStyle): \Vtiful\Kernel\Format
```

###example

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$data = [
    ['viest1', 21, 100, "A"],
    ['viest2', 20, 80, "B"],
    ['viest3', 22, 70, "C"],
];

$format = new \Vtiful\Kernel\Format($fileHandle);

// Create a border style
$borderStyle = $format
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->toResource();

$fileObject->header(['name', 'age', 'score', 'level'])
    ->data($data)
    ->setRow('A1', 20, $borderStyle)
    ->output();
```