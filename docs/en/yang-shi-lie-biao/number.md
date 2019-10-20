# Number format

## **Function Prototype**

```php
Number(string $format): self
```

### **string $format**

> array format string

```
"0.000"
"#,##0"
"#,##0.00"
"0.00"
```

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
$numberStyle = $format->number('#,##0')->toResource();

$filePath = $fileObject->header(['name', 'balance'])
     ->data([
         ['viest', 10000],
         ['wjx', 100000]
     ])
     ->setColumn('B:B', 50, $numberStyle) // Apply style
     ->output();
```