# Font

## **Function Prototype**

```php
Font(string $fontName): self
```

### **string $fontName**

> font name, font must exist in this machine

## Example

```php
$format = new \Vtiful\Kernel\Format($fileHandle);
$fontStyle = $format->font('FontName')->toResource();
```