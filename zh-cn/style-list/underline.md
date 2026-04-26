# 下划线

## **函数原型**

```php
underline(resource $resourchHandle, Format::const $style): \Vtiful\Kernel\Format
```

## 示例

```php
$format         = new \Vtiful\Kernel\Format($fileHandle);
$underlineStyle = $format->underline(Format::UNDERLINE_SINGLE)->toResource();
```

## **Style**

```php
Format::UNDERLINE_SINGLE;            // 单下划线
Format::UNDERLINE_DOUBLE;            // 双下划线
Format::UNDERLINE_SINGLE_ACCOUNTING; // 会计用单下划线
Format::UNDERLINE_DOUBLE_ACCOUNTING; // 会计用双下划线
```

