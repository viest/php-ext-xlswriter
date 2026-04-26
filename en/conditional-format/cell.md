# Single cell rule

Apply a conditional format rule to a single cell.

## Methods

```php
conditionalFormatCell(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
```

### **string $rangeA1**

> Target cell address in A1 notation, e.g. `"A1"` or `"C5"`.

### **\Vtiful\Kernel\ConditionalFormat $cf**

> A rule object configured via `type()` / `criteria()` / `value()` / `format()` etc.

## Greater-than

Highlight `A2` with a red background and white font when its value is greater than `50`:

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
    ->insertText(1, 0, 80) // A2 = 80, fires the rule
    ->conditionalFormatCell('A2', $cf)
    ->output();
```

## Between

Highlight `B2` when its value is between `60` and `90` (inclusive):

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$highlight = (new \Vtiful\Kernel\Format($fileHandle))
    ->background(\Vtiful\Kernel\Format::COLOR_YELLOW)
    ->toResource();

$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_BETWEEN)
   ->minimum(60)
   ->maximum(90)
   ->format($highlight);

$fileObject->header(['name', 'score'])
    ->data([['viest', 75]])
    ->conditionalFormatCell('B2', $cf)
    ->output();
```

## Text contains

Highlight `A2` when its text contains `error`:

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_TEXT)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_TEXT_CONTAINING)
   ->valueString('error')
   ->format($highlight);

$fileObject->conditionalFormatCell('A2', $cf);
```
