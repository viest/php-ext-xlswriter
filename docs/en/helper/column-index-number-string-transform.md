# Column index transform

### Number column index to string

```php
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(0));  // A
var_dump(\Vtiful\Kernel\Excel::stringFromColumnIndex(28)); // AC
```

### String column index to number

```php
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('A'));  // 0
var_dump(\Vtiful\Kernel\Excel::columnIndexFromString('AC')); // 28
```