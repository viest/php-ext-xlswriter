# 标签颜色

## **函数原型**

```php
setTabColor(int $rgb): self
```

### **int $rgb**

> 工作表标签颜色，使用 `0xRRGGBB` 整数表达，也可以直接使用 `Format::COLOR_*` 常量。
> 例如 `0xFF0000` 为红色，`0x00B050` 为绿色。

## 示例

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
