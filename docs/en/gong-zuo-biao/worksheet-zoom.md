# Scaling ratio

## **Function Prototype**

```php
zoom(int $scale = 100): self
```

### **int $scale**

> sheet zoom
>
> Range: 10 <= $scale <= 400
>
> Default: 100
>
> Scale does not affect print scale

## **Instance**

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
     ->zoom(200) // Set the worksheet zoom factor
     ->data([
     ['viest', 21],
     ['viest', 22],
     ['viest', 23],
     ])
     ->output();
```