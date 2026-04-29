# Row cell style

### **Function Prototype**

```php
setRow(string $range, double $height [, resource $formatHandler]);
```

#### **string $range**

> Cell range

#### **double $height**

> cell height

#### **resource $formatHandler**

> cell style

###example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->header(['name', 'age'])
     ->data([['viest', 21]])
     ->setRow('A1', 20, $boldStyle)
     ->output();
```