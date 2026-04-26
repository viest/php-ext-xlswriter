# 区域规则

将一条条件格式规则应用到一个 A1 区域。区域内每个单元格都会独立判断规则。

## 函数原型

```php
conditionalFormatRange(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
```

### **string $rangeA1**

> 目标区域，使用 A1 表示法，例如 `"A2:A10"`、`"B2:D20"`。

### **\Vtiful\Kernel\ConditionalFormat $cf**

> 通过 `\Vtiful\Kernel\ConditionalFormat` 构造的规则对象。

## 单元格大于约束

将 `A2:A10` 中所有大于 `50` 的单元格标记为红底白字：

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
   ->format($highlight);

$fileObject->header(['score'])
    ->data([[10], [40], [55], [60], [70], [80], [90], [100], [120]])
    ->conditionalFormatRange('A2:A10', $cf)
    ->output();
```

## 双色色阶

按数值大小在 `A2:A10` 上绘制蓝-红渐变：

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_2_COLOR_SCALE)
   ->minimumRule(\Vtiful\Kernel\ConditionalFormat::RULE_MINIMUM)
   ->minimumColor(0xFFFFFF) // 白色
   ->maximumRule(\Vtiful\Kernel\ConditionalFormat::RULE_MAXIMUM)
   ->maximumColor(0x63BE7B); // 绿色

$fileObject->conditionalFormatRange('A2:A10', $cf);
```

## 三色色阶

低-中-高三色渐变（红-黄-绿）：

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_3_COLOR_SCALE)
   ->minimumRule(\Vtiful\Kernel\ConditionalFormat::RULE_MINIMUM)
   ->minimumColor(0xF8696B)
   ->middleRule(\Vtiful\Kernel\ConditionalFormat::RULE_PERCENTILE)
   ->middle(50)
   ->middleColor(0xFFEB84)
   ->maximumRule(\Vtiful\Kernel\ConditionalFormat::RULE_MAXIMUM)
   ->maximumColor(0x63BE7B);

$fileObject->conditionalFormatRange('A2:A10', $cf);
```

## 公式约束

通过任意 Excel 公式判断（例如：偶数行高亮）：

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_FORMULA)
   ->valueString('=MOD(ROW(),2)=0')
   ->format($highlight);

$fileObject->conditionalFormatRange('A2:A10', $cf);
```
