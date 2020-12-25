# Print orientation-landscape

## **Function Prototype**

```php
setPrintedLandscape(): self
```

## **Example**

```php
$config = ['path' =>'./tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_landscape.xlsx','sheet1')
    ->setPrintedLandscape() // Set the printing direction to landscape
    ->output();
```