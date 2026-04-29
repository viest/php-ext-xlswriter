# Data bar

A data bar draws a horizontal coloured bar inside the cell whose length is proportional to the value. It is the easiest way to compare the magnitudes in a column at a glance.

## Methods

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

> Colour value as a `0xRRGGBB` integer, or one of the `\Vtiful\Kernel\Format::COLOR_*` constants.

### **bool $on**

> Toggle, defaults to `true`. `barOnly` hides the cell value and shows only the bar; `barSolid` paints a flat fill (Excel 2010 style); `dataBar2010` enables the Excel 2010 extension attributes; `barNoBorder` removes the bar border.

### **int $direction**

> Bar direction: `BAR_DIRECTION_CONTEXT` / `BAR_DIRECTION_LEFT_TO_RIGHT` / `BAR_DIRECTION_RIGHT_TO_LEFT`.

### **int $position**

> Axis position: `BAR_AXIS_AUTOMATIC` / `BAR_AXIS_MIDPOINT` / `BAR_AXIS_NONE`.

## Basic data bar

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
   ->barColor(0x638EC6); // blue

$excel->fileName('tutorial.xlsx')
    ->header(['score'])
    ->data([[10], [40], [55], [60], [70], [80], [90], [100]])
    ->conditionalFormatRange('A2:A9', $cf)
    ->output();
```

## Data bar with positive and negative values

Enable the Excel 2010 extension so negative values use a different colour, and place the axis at the midpoint:

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
   ->dataBar2010()
   ->barColor(0x63BE7B)            // green for positive
   ->barNegativeColor(0xF8696B)    // red for negative
   ->barAxisPosition(\Vtiful\Kernel\ConditionalFormat::BAR_AXIS_MIDPOINT)
   ->barAxisColor(0x000000)
   ->barSolid();

$excel->fileName('tutorial.xlsx')
    ->header(['delta'])
    ->data([[-30], [-10], [0], [20], [50], [70]])
    ->conditionalFormatRange('A2:A7', $cf)
    ->output();
```

## Bar only (hide the number)

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_DATA_BAR)
   ->barColor(0x638EC6)
   ->barOnly()       // hide the cell value
   ->barNoBorder();  // no border around the bar

$excel->conditionalFormatRange('A2:A9', $cf);
```
