# 锁定保护

启用工作表保护后，用户在 Excel 中将无法编辑被保护的单元格。保护可以选择是否带密码，也可以与 `Format::unlocked()` 配合，单独放开特定行、列或单元格的编辑权限。

## 函数原型

```php
protection(?string $password = null): self
```

### **string $password**

> 可选参数，解除保护所需的密码。省略时工作表仍然处于保护状态，但任何人都可以在 Excel 界面中无密码地解除保护。

## 子章节

* [密码保护](password.md)
* [解除保护](unlock.md)

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('tutorial.xlsx')
    ->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->protection('viest') // 设置密码为 viest
    ->output();
```
