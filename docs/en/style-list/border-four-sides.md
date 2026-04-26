# Border of the four sides

Apply a different border style to each of the cell's four sides.

## **Function Prototype**

```php
borderOfTheFourSides(
    int $top    = Format::BORDER_NONE,
    int $right  = Format::BORDER_NONE,
    int $bottom = Format::BORDER_NONE,
    int $left   = Format::BORDER_NONE
): self
```

### **int $top / $right / $bottom / $left**

> Border style of each side, taken from the `Format::BORDER_*` constants (e.g. `BORDER_THIN`, `BORDER_MEDIUM`, `BORDER_DASHED`).
> Any argument that is omitted or `null` leaves that side blank.

Pair this with `borderColorOfTheFourSides()` when you also want each side painted in a different color.

## Example

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$borderStyle = (new \Vtiful\Kernel\Format($fileHandle))
    ->borderOfTheFourSides(
        \Vtiful\Kernel\Format::BORDER_THIN,    // top
        \Vtiful\Kernel\Format::BORDER_MEDIUM,  // right
        \Vtiful\Kernel\Format::BORDER_THIN,    // bottom
        \Vtiful\Kernel\Format::BORDER_DASHED   // left
    )
    ->toResource();

$fileObject->header(['name', 'score'])
    ->setRow('A1', 20, $borderStyle)
    ->output();
```
