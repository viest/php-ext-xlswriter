# 多区域

一条条件格式规则可以同时应用到多个不连续的区域，并通过 `stopIfTrue` 控制是否在命中后跳过其它规则。

## 函数原型

```php
ConditionalFormat::multiRange(string $range): self
ConditionalFormat::stopIfTrue(bool $on = true): self
```

### **string $range**

> 多个 A1 区域，使用空格分隔，例如 `"A1:A5 C1:C5 E1:E10"`。该字段会替换 `conditionalFormatRange` 时使用的连续区域，让一条规则覆盖更多目标。

### **bool $on**

> `stopIfTrue` 是否启用「命中后停止」语义，默认 `true`。当多条规则都覆盖同一个单元格时，启用后只要前一条命中，后续规则将不会再被求值。

## 多区域示例

将「值大于 50 高亮」规则同时应用到 `A2:A6`、`C2:C6`、`E2:E6` 三个不连续区域：

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

// conditionalFormatRange 的第一个参数仍要写一个起始区域，
// 其它区域以 multiRange() 中的字符串为准。
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

## stopIfTrue 示例

当一条规则需要优先于另一条规则被求值时使用：

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

上例中，大于 90 的单元格只会应用金色样式（`stopIfTrue` 阻止了第二条规则继续生效），介于 60 与 90 之间的单元格则只会应用第二条黄色样式。
