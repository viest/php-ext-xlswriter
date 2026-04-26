# Merge Cells

### **Function Prototype**

```php
mergeCells(string $scope, string $data[, resource $formatHandler]): self
```

#### **string $scope**

> Cell range

#### **string $data**

> Data

#### **resource $formatHandler**

> cell style

###example

```php
$excel->fileName("test.xlsx")
   ->mergeCells('A1:C1', 'Merge cells')
   ->output();
```