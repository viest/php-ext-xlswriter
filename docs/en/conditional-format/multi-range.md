# Multi range

A single conditional format rule may target several non-contiguous ranges. `stopIfTrue` controls whether subsequent rules are evaluated after this rule fires.

## Methods

```php
ConditionalFormat::multiRange(string $range): self
ConditionalFormat::stopIfTrue(bool $on = true): self
```

### **string $range**

> Multiple A1 ranges separated by spaces, e.g. `"A1:A5 C1:C5 E1:E10"`. This list replaces the contiguous range used by `conditionalFormatRange`, letting one rule cover several disjoint targets.

### **bool $on**

> Whether `stopIfTrue` is enabled, defaults to `true`. When several rules cover the same cell, enabling this prevents later rules from being evaluated once this rule matches.

## Multi-range example

Apply the "highlight values greater than 50" rule to three disjoint ranges `A2:A6`, `C2:C6`, and `E2:E6`:

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$highlight = (new \Vtiful\Kernel\Format($fileHandle))
    ->background(\Vtiful\Kernel\Format::COLOR_RED)
    ->fontColor(\Vtiful\Kernel\Format::COLOR_WHITE)
    ->toResource();

$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
   ->value(50)
   ->format($highlight)
   ->multiRange('A2:A6 C2:C6 E2:E6');

// conditionalFormatRange still needs a starting range for the first parameter;
// the actual targets come from the multiRange() string above.
$fileObject->header(['x', '_', 'y', '_', 'z'])
    ->data([
        [10, '', 60, '', 30],
        [70, '', 20, '', 80],
        [55, '', 45, '', 90],
        [40, '', 35, '', 25],
        [85, '', 65, '', 15],
    ])
    ->conditionalFormatRange('A2:A6', $cf)
    ->output();
```

## stopIfTrue example

Use this when one rule must take precedence over another:

```php
$first = new \Vtiful\Kernel\ConditionalFormat();
$first->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
      ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
      ->value(90)
      ->format($highlightGold)
      ->stopIfTrue();

$second = new \Vtiful\Kernel\ConditionalFormat();
$second->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
       ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
       ->value(60)
       ->format($highlightYellow);

$fileObject->conditionalFormatRange('A2:A10', $first)
           ->conditionalFormatRange('A2:A10', $second);
```

In this example cells greater than 90 receive only the gold style (because `stopIfTrue` blocks the second rule), while cells between 60 and 90 receive only the yellow one.
