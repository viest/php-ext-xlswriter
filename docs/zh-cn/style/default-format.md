# 全局默认样式

设置全局默认样式，将会影响所有写入单元格的样式；

## **函数原型**

```php
defaultFormat(resource $formatHandler)
```

### **resource $formatHandler**

> 单元格样式

## 示例

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

