# Merge Cells

### **Function Prototype**

```php
mergeCells(string $scope, string $data): self
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