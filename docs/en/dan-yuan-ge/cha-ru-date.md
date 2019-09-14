# Insert date

## **Function Prototype**

```php
insertText(int $row, int $column, int $timestamp[, string $dateFormat = 'yyyy-mm-dd hh:mm:ss', resource $style])
```

### **int $row**

> cell row

### **int $column**

> cell column

### **int $timestamp**

> Timestamp to write

### **string $dataFormat**

> time formatting characters
>
> Default: yyyy-mm-dd hh:mm:ss

### **resource $style**

> cell style

##example

```php
$excel = new \Vtiful\Kernel\Excel($config);

$dateFile = $excel->fileName("free.xlsx")
     ->header(['date']);

$dateFile->insertDate(1, 0, time(), 'mmm d yyyy hh:mm AM/PM');

$textFile->output();
```

## Time Format Character Example

For more styles, please refer to the Excel Microsoft documentation.

```php
"mm/dd/yy"
"mmm d yyyy"
"d mmmm yyyy"
"dd/mm/yyyy hh:mm AM/PM"
```