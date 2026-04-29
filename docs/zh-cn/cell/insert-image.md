# 插入本地图片

## **函数原型**

```php
insertImage(int $row, int $column, string $localImagePath[, double $widthScale, double $heightScale])
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **string $localImagePath**

> 图片路径

### **double $widthScale**

> 对图像X轴进行缩放处理； 默认为1，保持图像原始宽度；值为0.5时，图像宽度为原图的1/2；

### **double $heightScale**

> 对图像轴进行缩放处理； 默认为1，保持图像原始高度；值为0.5时，图像高度为原图的1/2；

## 示例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("free.xlsx");

$freeFile->insertImage(5, 0, '/vagrant/ASW-G-66.jpg');

$freeFile->output();
```

