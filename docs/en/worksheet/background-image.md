# Background image

Set a tiled background image for the current worksheet. Use the path-based variant when the image is on disk, or the buffer-based variant when the bytes already live in memory (streaming, S3, …).

## **Function Prototype**

```php
setBackgroundImage(string $imagePath): self

setBackgroundImageBuffer(string $imageBuffer): self
```

### **string $imagePath**

> Path to a PNG or JPEG file.

### **string $imageBuffer**

> Raw image bytes, e.g. the return value of `file_get_contents()`.

## Example

```php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('tutorial.xlsx', 'sheet1')
      ->setBackgroundImage('./assets/logo.png')
      ->output();
```

```php
$excel = new \Vtiful\Kernel\Excel($config);

$buffer = file_get_contents('./assets/logo.png');

$excel->fileName('tutorial.xlsx', 'sheet1')
      ->setBackgroundImageBuffer($buffer)
      ->output();
```
