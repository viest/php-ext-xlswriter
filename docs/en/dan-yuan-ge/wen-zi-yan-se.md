# Text color

### **Function Prototype**

```php
fontColor(int $color)
```

#### **int $color**

> RGB hexadecimal value

###example

```php
$config     = ['path' => './tests'];
$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->fontColor(0xFF0000)->toResource();
// or
// $colorStyle = $format->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$filePath = $fileObject->header(['name', 'age'])
     ->data([
         ['viest', 21],
         ['wjx', 21]
     ])
     ->setRow('A1', 50, $colorStyle)
     ->output();

var_dump($filePath);
```