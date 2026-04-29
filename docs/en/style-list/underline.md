# Underline

### **Function Prototype**

```php
Underline(resource $resourchHandle, Format::const $style): \Vtiful\Kernel\Format
```

### Example

```php
$format         = new \Vtiful\Kernel\Format($fileHandle);
$underlineStyle = $format->underline(Format::UNDERLINE_SINGLE)->toResource();
```

### **Style**

```php
Format::UNDERLINE_SINGLE;            // Single underline
Format::UNDERLINE_DOUBLE;            // double underline
Format::UNDERLINE_SINGLE_ACCOUNTING; // Accounting underline
Format::UNDERLINE_DOUBLE_ACCOUNTING; // Accounting double underline
```