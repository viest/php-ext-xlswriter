# 插入公式

## **函数原型**

```php
insertFormula(int $row, int $column, string $formula [, resource $formatHandler])
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **string $formula**

> 公式

### **resource $formatHandler**

> 单元格样式

## 示例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$freeFile = $excel->fileName("free.xlsx")
    ->header(['name', 'money']);

for($index = 1; $index < 10; $index++) {
    $freeFile->insertText($index, 0, 'viest');
    $freeFile->insertText($index, 1, 10);
}

$freeFile->insertText(12, 0, "Total");
$freeFile->insertFormula(12, 1, '=SUM(B2:B11)');

$freeFile->output();
```

