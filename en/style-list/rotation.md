# Rotation

Rotate the text inside a cell.

## **Function Prototype**

```php
rotation(int $angle): self
```

### **int $angle**

> Angle in degrees.
>
> - A positive value rotates the text counter-clockwise (e.g. `45` tilts up-and-to-the-left).
> - A negative value rotates clockwise (e.g. `-45` tilts down-and-to-the-right).
> - Valid range is `-90..90`.
> - The special value `270` lays the text out vertically (each character stacked beneath the previous one).
>
> Values outside the valid range are silently ignored.

## Example

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$rotate45 = (new \Vtiful\Kernel\Format($fileHandle))
    ->rotation(45)
    ->toResource();

$vertical = (new \Vtiful\Kernel\Format($fileHandle))
    ->rotation(270)
    ->toResource();

$fileObject->header(['flat', 'tilted', 'vertical'])
    ->insertText(1, 0, 'normal')
    ->insertText(1, 1, 'tilted',   null, $rotate45)
    ->insertText(1, 2, 'vertical', null, $vertical)
    ->output();
```
