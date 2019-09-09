# 合并单元格

### **函数原型**

```text
mergeCells(string $scope, string $data);
```

#### **string $scope**

> 单元格范围

#### **string $data**

> 数据

### 示例

```php
$excel->fileName("test.xlsx")
  ->mergeCells('A1:C1', 'Merge cells')
  ->output();
```

