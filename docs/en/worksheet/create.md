# Create a worksheet

### **Function Prototype**

```php
addSheet([string $sheetName]);
```

### Example

```php
$config = [
     'path' => './filePath'
];

$excel = new \Vtiful\Kernel\Excel($config);

// A worksheet is automatically created here
$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
     ->data([['viest', 21]]);

// append a worksheet to the file
$fileObject->addSheet()
     ->header(['name', 'age'])
     ->data([['wjx', 22]]);

// Finally, the output file
$filePath = $fileObject->output();
```