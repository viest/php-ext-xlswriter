# 边框颜色

## **函数原型**

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

> 同时设置上下左右四条边框的颜色，使用 `0xRRGGBB` 整数或 `Format::COLOR_*` 常量。

### **int $topColor / $rightColor / $bottomColor / $leftColor**

> 分别设置上、右、下、左四条边框颜色。任意参数省略或传 `null` 时该方向保持默认色。

设置颜色之前应先通过 `border()` 或 `borderOfTheFourSides()` 指定边框样式，否则边框线本身不会显示。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// 四边同色
$singleColor = (new \Vtiful\Kernel\Format($fileHandle))
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->borderColor(\Vtiful\Kernel\Format::COLOR_RED)
    ->toResource();

// 四边不同色
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
