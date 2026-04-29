# 列标字符转化

### 数字转字符

```php
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(0));  // A
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(28)); // AC
```

### 字符转数字

```php
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('A'));  // 0
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('AC')); // 28
```