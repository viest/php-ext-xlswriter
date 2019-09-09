# 插入文字

## **函数原型**

```php
insertText(int $row, int $column, string|int|double $data[, string $format, resource $style])
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **string \| int \| double $data**

> 需要写入的内容

### **string $format**

> 内容格式

### **resource $style**

> 单元格样式

## 示例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$textFile = $excel->fileName("free.xlsx")
    ->header(['name', 'money']);

for ($index = 0; $index < 10; $index++) {
    $textFile->insertText($index+1, 0, 'viest');
    $textFile->insertText($index+1, 1, 10000, '#,##0'); // #,##0 为单元格数据样式
}

$textFile->output();
```

## 数字样式示例

更多样式请参考 Excel 微软文档

```php
"0.000"
"#,##0"
"#,##0.00"
"0.00"
"mm/dd/yy"
"mmm d yyyy"
"d mmmm yyyy"
"dd/mm/yyyy hh:mm AM/PM"
"0 \"dollar and\" .00 \"cents\""
```

