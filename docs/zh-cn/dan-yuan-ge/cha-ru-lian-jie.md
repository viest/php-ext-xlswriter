# 插入链接

## **函数原型**

```php
insertUrl(int $row, int $column, string $url[, resource $formatHandler])
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **string $url**

> 链接地址

### **resource $formatHandler**

> 链接样式

## 示例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$urlFile = $excel->fileName("free.xlsx")
    ->header(['url']);

$fileHandle = $urlFile->getHandle();

$format   = new \Vtiful\Kernel\Format($fileHandle);
$urlStyle = $format->bold()
    ->underline(\Vtiful\Kernel\Format::UNDERLINE_SINGLE)
    ->toResource();

$urlFile->insertUrl(1, 0, 'https://github.com', $urlStyle);

$urlFile->output();
```

