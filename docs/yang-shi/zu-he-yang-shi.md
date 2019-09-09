# 组合样式

将多个样式合并为一个新样式应用在单元格上

```php
// 将粗体与斜体合并为一个样式
$format          = new \Vtiful\Kernel\Format($fileHandle);
$boldItalicStyle = $format->bold()->italic()->toResour();
```

