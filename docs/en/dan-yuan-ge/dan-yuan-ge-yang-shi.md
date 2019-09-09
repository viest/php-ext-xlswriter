# Cell style

### **Function Prototype**

```php
setColumn(string $range, double $width [, resource $format]);
```

#### **string $range**

> Cell range

#### **double $width**

> cell width

#### **string $format**

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
     ->setColumn('A:A', 200, $boldStyle)
     ->output();
```