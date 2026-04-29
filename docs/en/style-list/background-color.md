# Background color

## **Function Prototype**

```php
background(int $color, int $pattern = self::PATTERN_SOLID): self
```

### **int $color**

> color const or RGB hex

### **int $pattern**

> pattern style

## Example

```php
$format = new \Vtiful\Kernel\Format($fileHandle);

$backgroundStyle  = $format->background(
   \Vtiful\Kernel\Format::COLOR_RED
)->toResource();
```

