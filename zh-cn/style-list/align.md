# 对齐

## **函数原型**

```php
align(resource $resourchHandle, Format::const ...$style): \Vtiful\Kernel\Format
```

## 示例

```php
$format     = new \Vtiful\Kernel\Format($fileHandle);
$alignStyle = $format
    ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
    ->toResource();
```

## **Style**

```php
Format::FORMAT_ALIGN_LEFT;                 // 水平左对齐
Format::FORMAT_ALIGN_CENTER;               // 水平剧中对齐
Format::FORMAT_ALIGN_RIGHT;                // 水平右对齐
Format::FORMAT_ALIGN_FILL;                 // 水平填充对齐
Format::FORMAT_ALIGN_JUSTIFY;              // 水平两端对齐
Format::FORMAT_ALIGN_CENTER_ACROSS;        // 横向中心对齐
Format::FORMAT_ALIGN_DISTRIBUTED;          // 分散对齐
Format::FORMAT_ALIGN_VERTICAL_TOP;         // 顶部垂直对齐
Format::FORMAT_ALIGN_VERTICAL_BOTTOM;      // 底部垂直对齐
Format::FORMAT_ALIGN_VERTICAL_CENTER;      // 垂直剧中对齐
Format::FORMAT_ALIGN_VERTICAL_JUSTIFY;     // 垂直两端对齐
Format::FORMAT_ALIGN_VERTICAL_DISTRIBUTED; // 垂直分散对齐
```

