# 合并单元格

## **函数原型**

```php
mergeCells(string $scope, string $data [, resource $formatHandler]): self
```

### **string $scope**

> 单元格范围

### **string $data**

> 数据

### **resource $formatHandler**

> 单元格样式

## 示例

### 横向合并单元格

```php
$excel->fileName("test.xlsx")
  ->mergeCells('A1:C1', 'Merge cells')
  ->output();
```

### 纵向合并单元格

```php
$excel->fileName("test.xlsx")
  ->mergeCells('A1:A3', 'Merge cells')
  ->output();
```
