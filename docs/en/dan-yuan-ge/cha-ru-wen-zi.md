# Insert text

### **Function Prototype**

```php
insertText(int $row, int $column, string|int|double $data[, string $format, resource $formatHandler])
```

#### **int $row**

> cell row

#### **int $column**

> cell column

#### **string \| int \| double $data**

> What needs to be written

#### **string $format**

> Content format

#### **resource $formatHandler**

> cell style

###example

```php
$excel = new \Vtiful\Kernel\Excel($config);

$textFile = $excel->fileName("free.xlsx")
     ->header(['name', 'money']);

for ($index = 0; $index < 10; $index++) {
     $textFile->insertText($index+1, 0, 'viest');
     $textFile->insertText($index+1, 1, 10000, '#,##0'); // #,##0 is the cell data style
}

$textFile->output();
```

### Digital Style Example

For more styles, please refer to the Excel Microsoft documentation.

```php
"0.000"
"#,##0"
"#,##0.00"
"0.00"
"0 \"dollar and\" .00 \"cents\""
```