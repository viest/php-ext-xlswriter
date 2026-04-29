# Icon set

An icon set draws a small icon to the left of the cell value based on the cell's relative position within the range. Common styles include three-colour traffic lights and five-star ratings.

## Methods

```php
ConditionalFormat::iconStyle(int $style): self
ConditionalFormat::reverseIcons(bool $on = true): self
ConditionalFormat::iconsOnly(bool $on = true): self
```

### **int $style**

> Icon style constant. See `\Vtiful\Kernel\ConditionalFormat::ICONS_3_*` / `ICONS_4_*` / `ICONS_5_*`.

### **bool $on**

> Toggle, defaults to `true`. `reverseIcons` flips the icon order (so high values use the "low" icon); `iconsOnly` hides the cell value and shows only the icon.

## Three-colour traffic lights

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

## Five-star rating (icons only)

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_ICON_SETS)
   ->iconStyle(\Vtiful\Kernel\ConditionalFormat::ICONS_5_RATINGS)
   ->iconsOnly();

$excel->conditionalFormatRange('A2:A9', $cf);
```

## Reversed icon order

For example, three coloured arrows: by default "higher = green"; with `reverseIcons` it becomes "lower = green":

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_ICON_SETS)
   ->iconStyle(\Vtiful\Kernel\ConditionalFormat::ICONS_3_ARROWS_COLORED)
   ->reverseIcons();

$excel->conditionalFormatRange('A2:A9', $cf);
```

## Icon style constants

All defined under `\Vtiful\Kernel\ConditionalFormat`:

```php
// Three icons
const ICONS_3_ARROWS_COLORED;
const ICONS_3_ARROWS_GRAY;
const ICONS_3_FLAGS;
const ICONS_3_TRAFFIC_LIGHTS_UNRIMMED;
const ICONS_3_TRAFFIC_LIGHTS_RIMMED;
const ICONS_3_SIGNS;
const ICONS_3_SYMBOLS_CIRCLED;
const ICONS_3_SYMBOLS_UNCIRCLED;

// Four icons
const ICONS_4_ARROWS_COLORED;
const ICONS_4_ARROWS_GRAY;
const ICONS_4_RED_TO_BLACK;
const ICONS_4_RATINGS;
const ICONS_4_TRAFFIC_LIGHTS;

// Five icons
const ICONS_5_ARROWS_COLORED;
const ICONS_5_ARROWS_GRAY;
const ICONS_5_RATINGS;
const ICONS_5_QUARTERS;
```
