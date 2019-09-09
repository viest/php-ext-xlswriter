# Background color

```php
$format = new \Vtiful\Kernel\Format($fileHandle);

$backgroundStyle  = $format->background(
   \Vtiful\Kernel\Format::PATTERN_LIGHT_UP,
   \Vtiful\Kernel\Format::COLOR_RED
)->toResource();
```

