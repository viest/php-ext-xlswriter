# 插入日期

## **函数原型**

```php
insertDate(int $row, int $column, int $timestamp[, string $dateFormat = 'yyyy-mm-dd hh:mm:ss', resource $style])
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **int $timestamp**

> 需要写入的时间戳

### **string $dataFormat**

> 时间格式化字符
>
> 默认值：yyyy-mm-dd hh:mm:ss

### **resource $style**

> 单元格样式

## 示例

```php
$excel = new \Vtiful\Kernel\Excel($config);

$dateFile = $excel->fileName("free.xlsx")
    ->header(['date']);

$dateFile->insertDate(1, 0, time(), 'mmm d yyyy hh:mm AM/PM');

$textFile->output();
```

## 时间格式化字符示例

更多样式请参考 Excel 微软文档

```php
"mm/dd/yy"
"mmm d yyyy"
"d mmmm yyyy"
"dd/mm/yyyy hh:mm AM/PM"
```

