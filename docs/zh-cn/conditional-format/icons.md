# 图标集

图标集（Icon Set）会根据单元格在区间内的相对位置在单元格左侧绘制一组图标，例如三色交通灯、五星评分等。

## 函数原型

```php
ConditionalFormat::iconStyle(int $style): self
ConditionalFormat::reverseIcons(bool $on = true): self
ConditionalFormat::iconsOnly(bool $on = true): self
```

### **int $style**

> 图标样式常量，参见 `\Vtiful\Kernel\ConditionalFormat::ICONS_3_*` / `ICONS_4_*` / `ICONS_5_*`。

### **bool $on**

> 开关参数，默认 `true`。`reverseIcons` 反转图标顺序（高分用低分图标）；`iconsOnly` 只显示图标，不显示单元格中的数字。

## 三色交通灯

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_ICON_SETS)
   ->iconStyle(\Vtiful\Kernel\ConditionalFormat::ICONS_3_TRAFFIC_LIGHTS_UNRIMMED);

$excel->fileName('tutorial.xlsx')
    ->header(['score'])
    ->data([[10], [40], [55], [60], [70], [80], [90], [100]])
    ->conditionalFormatRange('A2:A9', $cf)
    ->output();
```

## 五星评分（仅显示图标）

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_ICON_SETS)
   ->iconStyle(\Vtiful\Kernel\ConditionalFormat::ICONS_5_RATINGS)
   ->iconsOnly();

$excel->conditionalFormatRange('A2:A9', $cf);
```

## 反转图标方向

例如三色箭头：默认是「越大越绿」，启用 `reverseIcons` 后变为「越小越绿」：

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_ICON_SETS)
   ->iconStyle(\Vtiful\Kernel\ConditionalFormat::ICONS_3_ARROWS_COLORED)
   ->reverseIcons();

$excel->conditionalFormatRange('A2:A9', $cf);
```

## 图标样式常量

下方常量均位于 `\Vtiful\Kernel\ConditionalFormat`：

```php
// 三图标
const ICONS_3_ARROWS_COLORED;
const ICONS_3_ARROWS_GRAY;
const ICONS_3_FLAGS;
const ICONS_3_TRAFFIC_LIGHTS_UNRIMMED;
const ICONS_3_TRAFFIC_LIGHTS_RIMMED;
const ICONS_3_SIGNS;
const ICONS_3_SYMBOLS_CIRCLED;
const ICONS_3_SYMBOLS_UNCIRCLED;

// 四图标
const ICONS_4_ARROWS_COLORED;
const ICONS_4_ARROWS_GRAY;
const ICONS_4_RED_TO_BLACK;
const ICONS_4_RATINGS;
const ICONS_4_TRAFFIC_LIGHTS;

// 五图标
const ICONS_5_ARROWS_COLORED;
const ICONS_5_ARROWS_GRAY;
const ICONS_5_RATINGS;
const ICONS_5_QUARTERS;
```
