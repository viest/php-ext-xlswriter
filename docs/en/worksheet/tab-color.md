# Tab color

## **Function Prototype**

```php
setTabColor(int $rgb): self
```

### **int $rgb**

> Sheet tab color expressed as a `0xRRGGBB` integer. The `Format::COLOR_*` constants can be used as well.
> For example `0xFF0000` is red and `0x00B050` is green.

## Example

```php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('tutorial.xlsx', 'sheet1')
      ->setTabColor(0xFF0000)
      ->addSheet('sheet2')
      ->setTabColor(\Vtiful\Kernel\Format::COLOR_GREEN)
      ->output();
```
