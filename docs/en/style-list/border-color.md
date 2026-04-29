# Border color

## **Function Prototype**

```php
borderColor(int $color): self

borderColorOfTheFourSides(
    int $topColor    = -1,
    int $rightColor  = -1,
    int $bottomColor = -1,
    int $leftColor   = -1
): self
```

### **int $color**

> Set the color of all four border sides at once. Pass a `0xRRGGBB` integer or a `Format::COLOR_*` constant.

### **int $topColor / $rightColor / $bottomColor / $leftColor**

> Set the four sides individually. Any argument that is omitted or `null` keeps the default color for that side.

The border line itself must be configured first through `border()` or `borderOfTheFourSides()`; setting only the color will not draw any line.

## Example

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// Same color on every side.
$singleColor = (new \Vtiful\Kernel\Format($fileHandle))
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->borderColor(\Vtiful\Kernel\Format::COLOR_RED)
    ->toResource();

// Different color on each side.
$mixedColor = (new \Vtiful\Kernel\Format($fileHandle))
    ->borderOfTheFourSides(
        \Vtiful\Kernel\Format::BORDER_THIN,
        \Vtiful\Kernel\Format::BORDER_THIN,
        \Vtiful\Kernel\Format::BORDER_THIN,
        \Vtiful\Kernel\Format::BORDER_THIN
    )
    ->borderColorOfTheFourSides(
        \Vtiful\Kernel\Format::COLOR_RED,
        \Vtiful\Kernel\Format::COLOR_GREEN,
        \Vtiful\Kernel\Format::COLOR_BLUE,
        \Vtiful\Kernel\Format::COLOR_YELLOW
    )
    ->toResource();

$fileObject->header(['name', 'score'])
    ->setRow('A1', 20, $singleColor)
    ->setRow('A2', 20, $mixedColor)
    ->output();
```
