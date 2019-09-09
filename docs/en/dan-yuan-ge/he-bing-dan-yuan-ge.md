# Merge Cells

### **Function Prototype**

```text
mergeCells(string $scope, string $data);
```

#### **string $scope**

> Cell range

#### **string $data**

> Data

###example

```php
$excel->fileName("test.xlsx")
   ->mergeCells('A1:C1', 'Merge cells')
   ->output();
```