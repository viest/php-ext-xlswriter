# Range rule

Apply a conditional format rule to an A1 range. Each cell in the range is evaluated independently.

## Methods

```php
conditionalFormatRange(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
```

### **string $rangeA1**

> Target range in A1 notation, e.g. `"A2:A10"` or `"B2:D20"`.

### **\Vtiful\Kernel\ConditionalFormat $cf**

> A rule object built with `\Vtiful\Kernel\ConditionalFormat`.

## Greater-than across a range

Mark every cell in `A2:A10` greater than `50` with a red background and white font:

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

## Two-colour scale

Apply a white-to-green gradient across `A2:A10`:

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_2_COLOR_SCALE)
   ->minimumRule(\Vtiful\Kernel\ConditionalFormat::RULE_MINIMUM)
   ->minimumColor(0xFFFFFF) // white
   ->maximumRule(\Vtiful\Kernel\ConditionalFormat::RULE_MAXIMUM)
   ->maximumColor(0x63BE7B); // green

$fileObject->conditionalFormatRange('A2:A10', $cf);
```

## Three-colour scale

Low / mid / high red-yellow-green gradient:

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

## Formula rule

Match cells via an arbitrary Excel formula (here: highlight even rows):

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_FORMULA)
   ->valueString('=MOD(ROW(),2)=0')
   ->format($highlight);

$fileObject->conditionalFormatRange('A2:A10', $cf);
```
