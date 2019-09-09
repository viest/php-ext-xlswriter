# Combination style

Combine multiple styles into one new style applied to the cell

```php
// Combine bold and italic into one style
$format          = new \Vtiful\Kernel\Format($fileHandle);
$boldItalicStyle = $format->bold()->italic()->toResour();
```