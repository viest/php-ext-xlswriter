# 数据条

数据条（Data Bar）会在单元格内按数值大小绘制一条横向色条，是直观比较一列数值的常用方式。

## 函数原型

```php
ConditionalFormat::barColor(int $color): self
ConditionalFormat::barOnly(bool $on = true): self
ConditionalFormat::barSolid(bool $on = true): self
ConditionalFormat::dataBar2010(bool $on = true): self
ConditionalFormat::barNegativeColor(int $color): self
ConditionalFormat::barBorderColor(int $color): self
ConditionalFormat::barNegativeBorderColor(int $color): self
ConditionalFormat::barNoBorder(bool $on = true): self
ConditionalFormat::barDirection(int $direction): self
ConditionalFormat::barAxisPosition(int $position): self
ConditionalFormat::barAxisColor(int $color): self
```

### **int $color**

> 颜色值，使用 0xRRGGBB 形式的整型，或 `\Vtiful\Kernel\Format::COLOR_*` 常量。

### **bool $on**

> 开关参数，默认 `true`。`barOnly` 表示只显示数据条不显示数字；`barSolid` 表示使用纯色（Excel 2010 风格）；`dataBar2010` 表示启用 Excel 2010 扩展属性；`barNoBorder` 表示不绘制边框。

### **int $direction**

> 数据条方向，使用 `BAR_DIRECTION_CONTEXT` / `BAR_DIRECTION_LEFT_TO_RIGHT` / `BAR_DIRECTION_RIGHT_TO_LEFT`。

### **int $position**

> 轴位置，使用 `BAR_AXIS_AUTOMATIC` / `BAR_AXIS_MIDPOINT` / `BAR_AXIS_NONE`。

## 基础数据条

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
   ->barColor(0x638EC6); // 蓝色

$excel->fileName('tutorial.xlsx')
    ->header(['score'])
    ->data([[10], [40], [55], [60], [70], [80], [90], [100]])
    ->conditionalFormatRange('A2:A9', $cf)
    ->output();
```

## 含正负值的数据条

启用 Excel 2010 扩展，使负数使用不同颜色，并把轴线放在中点：

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
   ->dataBar2010()
   ->barColor(0x63BE7B)            // 正值绿色
   ->barNegativeColor(0xF8696B)    // 负值红色
   ->barAxisPosition(\Vtiful\Kernel\ConditionalFormat::BAR_AXIS_MIDPOINT)
   ->barAxisColor(0x000000)
   ->barSolid();

$excel->fileName('tutorial.xlsx')
    ->header(['delta'])
    ->data([[-30], [-10], [0], [20], [50], [70]])
    ->conditionalFormatRange('A2:A7', $cf)
    ->output();
```

## 仅显示数据条

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
   ->barColor(0x638EC6)
   ->barOnly()       // 隐藏单元格里的数字
   ->barNoBorder();  // 不绘制边框

$excel->conditionalFormatRange('A2:A9', $cf);
```
