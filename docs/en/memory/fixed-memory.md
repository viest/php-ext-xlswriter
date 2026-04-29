# Fixed memory mode

### **Memory**

Maximum memory usage = maximum one row of data usage

### **Function Prototype**

```php
constMemory(string $fileName);
```

### Example

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->constMemory('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->header(['name', 'age'])
     ->data([['viest', 21]])
     ->setRow('A1', 10, $boldStyle)
     ->output();
```