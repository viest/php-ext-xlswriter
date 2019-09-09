# 文本换行

如果单元格内文本包含 `\n` ，将处理换行样式。

```php
$format    = new \Vtiful\Kernel\Format($fileHandle);
$wrapStyle = $format->wrap()->toResource();
```



