# Text color

## **Function Prototype**

```php
fontColor(int $color): self
```

### **int $color**

> RGB hexadecimal value or color constant

## Example

```php
$config = [
     'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// Create a style resource
$format = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$filePath = $fileObject->header(['name', 'age'])
     ->data([
         ['viest', 21],
         ['wjx', 21]
     ])
     ->setRow('A1', 50, $colorStyle) // Apply style
     ->output();
```