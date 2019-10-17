# 合并单元格

## **函数原型**

```php
mergeCells(string $scope, string $data): self
```

### **string $scope**

> 单元格范围

### **string $data**

> 数据

## 示例

```php
$excel->fileName("test.xlsx")
  ->mergeCells('A1:C1', 'Merge cells')
  ->output();
```

