# Insert local image

### **Function Prototype**

```php
insertImage(int $row, int $column, string $localImagePath[, double $widthScale, double $heightScale])
```

#### **int $row**

> cell row

#### **int $column**

> cell column

#### **string $localImagePath**

> picture path

#### **double $widthScale**

> Scale the image X axis; the default is 1, maintaining the original width of the image; when the value is 0.5, the image width is 1/2 of the original image;

#### **double $heightScale**

> Scale the image axis; the default is 1, keeping the original height of the image; when the value is 0.5, the image height is 1/2 of the original image;

###example

```php
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("free.xlsx");

$freeFile->insertImage(5, 0, '/vagrant/ASW-G-66.jpg');

$freeFile->output();
```