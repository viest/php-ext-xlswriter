# Default Format

Setting the global default style will affect the style of all written cells;

## **Function prototype**

```php
defaultFormat(resource $formatHandler)
```

### **resource $formatHandler**

> cell style

## Example

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('tutorial.xlsx');

$format        = new \Vtiful\Kernel\Format($excel->getHandle());
$colorOneStyle = $format
    ->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)
    ->border(\Vtiful\Kernel\Format::BORDER_DASH_DOT)
    ->toResource();

$format        = new \Vtiful\Kernel\Format($excel->getHandle());
$colorTwoStyle = $format
    ->fontColor(\Vtiful\Kernel\Format::COLOR_GREEN)
    ->toResource();

$filePath = $excel
    // Apply the first style as the default
    ->defaultFormat($colorOneStyle)
    ->header(['hello', 'xlswriter'])
    // Apply the second style as the default style
    ->defaultFormat($colorTwoStyle)
    ->data([
        ['hello', 'xlswriter'],
    ])
    ->output();
```

