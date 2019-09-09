# 单元格样式

## **函数原型**

```php
setColumn(string $range, double $width [, resource $format]);
```

### **string $range**

> 单元格范围

### **double $width**

> 单元格宽度

### **string $format**

> 单元格样式

## 示例

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

