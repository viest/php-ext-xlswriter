# Switch worksheet

### **Function Prototype**

```php
checkoutSheet(string $sheetName);
```

### **Instance**

```php
$config = [
   'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
     ->data([
     ['viest', 21],
     ['viest', 22],
     ['viest', 23],
     ]);

// Add a worksheet and insert data
$fileObject->addSheet('twoSheet')
     ->header(['name', 'age'])
     ->data([['vikin', 22]]);

/ / Switch back to the default work table, and append data
$fileObject->checkoutSheet('Sheet1')
     ->data([['sheet1']]);

$filePath = $fileObject->output();
```