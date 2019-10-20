# Gridlines

## **Function Prototype**

```php
gridline(int $option = \Vtiful\Kernel\Excel::GRIDLINES_SHOW_ALL): self
```

## Grid line type

```php
Const GRIDLINES_HIDE_ALL    = 0; // hide screen grid lines and print grid lines
Const GRIDLINES_SHOW_SCREEN = 1; // Display screen grid lines
Const GRIDLINES_SHOW_PRINT  = 2; // Display print grid lines
Const GRIDLINES_SHOW_ALL    = 3; // display screen grid lines and print grid lines
```

## **Instance**

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
     ->gridline(\Vtiful\Kernel\Excel::GRIDLINES_HIDE_ALL) // Set the worksheet grid lines
     ->data([
         ['viest', 21],
         ['viest', 22],
         ['viest', 23],
     ])
     ->output();
```