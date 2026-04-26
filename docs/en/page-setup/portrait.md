# Print direction-portrait

## **Function Prototype**

```php
setPortrait(): self
```

## **Example**

```php
$config = ['path' =>'./tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$excel->fileName('printed_portrait.xlsx','sheet1')
    ->setPortrait() // Set the printing direction to portrait
    ->output();
```
