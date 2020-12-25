# Set the current worksheet as the first worksheet

## **Function Prototype**

```php
setCurrentSheetIsFirst(): self
```

## **Example**

```php
$config = ['path' =>'./tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('hide.xlsx','sheet1') // Initialize the file and initialize the first sheet at the same time sheet1
    ->header(['sheet1'])       // insert data in sheet1 worksheet
    ->addSheet('sheet2')       // Add a new sheet sheet2, and set the current active sheet to sheet2
    ->setCurrentSheetIsFirst() // The current active sheet is sheet2, and set sheet2 as the first sheet
    ->output();
```