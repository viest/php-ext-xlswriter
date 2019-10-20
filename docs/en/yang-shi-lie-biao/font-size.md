# Text size

## **Function Prototype**

```php
fontSize(double $size);
```

### **double $size**

> cell font size

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
$style = $format->fontSize(30)->toResource();

$filePath = $fileObject->header(['name', 'age'])
     ->data([
         ['viest', 21],
         ['wjx', 21]
     ])
     ->setRow('A1', 50, $style) // Apply style
     ->setRow('A2:A3', 50, $style) // Apply style
     ->output();
```