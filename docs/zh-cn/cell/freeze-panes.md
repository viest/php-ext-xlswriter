# 冻结单元格

## **函数原型**

```php
freezePanes(int $row, int $column): self
```

### **int $row**

> 行编号

### **int $column**

> 列编号

## 示例

```php
freezePanes(1, 0); // 冻结第一行
freezePanes(0, 1); // 冻结第一列
freezePanes(1, 1); // 冻结第一行和第一列
```

