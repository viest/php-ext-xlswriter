# 行单元格样式

`setRow` 样式影响范围为整行。

设置 `range` 参数为 `A1:D1`，第一反应是设置第一行的前四个单元格样式，但是实际效果确是设置 `第一行整行`。

## **函数原型**

```php
setRow(string $range, double $height [, resource $formatHandler]);
```

### **string $range**

> 单元格范围

### **double $height**

> 单元格高度

### **resource $formatHandler**

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
    ->setRow('A1', 20, $boldStyle)
    ->output();
```

