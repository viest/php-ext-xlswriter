# 行单元格样式

## **函数原型**

```php
setRow(string $range, double $height [, resource $format]);
```

### **string $range**

> 单元格范围

### **double $height**

> 单元格高度

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
    ->setRow('A1', 20, $boldStyle,)
    ->output();
```

