# 插入链接

## **函数原型**

```php
insertUrl(int $row, int $column, string $url[, resource $format])
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **string $url**

> 链接地址

### **resource $format**

> 链接样式

## 示例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$urlFile = $excel->fileName("free.xlsx")
    ->header(['url']);

$fileHandle = $fileObject->getHandle();

$format   = new \Vtiful\Kernel\Format($fileHandle);
$urlStyle = $format->bold()
    ->underline(Format::UNDERLINE_SINGLE)
    ->toResource();

$urlFile->insertUrl(1, 0, 'https://github.com', $urlStyle);

$textFile->output();
```

