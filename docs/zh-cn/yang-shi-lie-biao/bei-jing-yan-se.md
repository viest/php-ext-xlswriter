# 背景颜色

## **函数原型**

```php
background(int $color, int $pattern = self::PATTERN_SOLID): self
```

### **int $color**

> 颜色常量

### **int $pattern**

> 背景图案样式，默认为实体背景

## Example

```php
$format = new \Vtiful\Kernel\Format($fileHandle);

$backgroundStyle  = $format->background(
   \Vtiful\Kernel\Format::COLOR_RED
)->toResource();
```

