# 列单元格样式

`setColumn` 样式影响范围为整列。

设置 `range` 参数为 `A1:D1`，第一反应是设置第一行的前四个单元格样式，但是实际效果确是设置 `第一列、第二列、第三列、第四列 整列`。

## **函数原型**

```php
setColumn(string $range, double $width [, resource $formatHandler]);
```

### **string $range**

> 单元格范围

### **double $width**

> 单元格宽度

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
    ->setColumn('A:A', 200, $boldStyle)
    ->output();
```
