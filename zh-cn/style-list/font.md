# 字体

## **函数原型**

```php
font(string $fontName): self
```

### **string $fontName**

> 字体名称，字体必须存在于本机

## 示例

```php
$format    = new \Vtiful\Kernel\Format($fileHandle);
$fontStyle = $format->font('FontName')->toResource();
```

