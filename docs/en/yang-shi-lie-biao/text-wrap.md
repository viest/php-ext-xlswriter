# Text wrap

If the text inside the cell contains `\n` , the newline style will be processed.

```php
$format    = new \Vtiful\Kernel\Format($fileHandle);
$wrapStyle = $format->wrap()->toResource();
```