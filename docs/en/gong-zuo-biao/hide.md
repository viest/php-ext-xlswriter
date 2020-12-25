# Hide the current worksheet

## **Function Prototype**

```php
setCurrentSheetHide(): self
```

## **Example**

```php
$config = ['path' =>'./tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('hide.xlsx','sheet1') // Initialize the file and initialize the first sheet at the same time sheet1
    ->header(['sheet1'])    // insert data in sheet1 worksheet
    ->addSheet('sheet2')    // Add a new sheet sheet2, and set the current active sheet to sheet2
    ->setCurrentSheetHide() // The current active sheet is sheet2, and sheet2 is hidden
    ->output();
```