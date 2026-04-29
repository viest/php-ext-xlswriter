# 单元格规则

将一条条件格式规则应用到单个单元格。

## 函数原型

```php
conditionalFormatCell(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
```

### **string $rangeA1**

> 目标单元格地址，A1 表示法，例如 `"A1"`、`"C5"`。

### **\Vtiful\Kernel\ConditionalFormat $cf**

> 已经通过 `type()` / `criteria()` / `value()` / `format()` 等方法配置好的规则对象。

## 大于约束

当单元格 `A1` 的值大于 `50` 时，应用红底白字样式：

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
    ->insertText(1, 0, 80) // A2 = 80，触发规则
    ->conditionalFormatCell('A2', $cf)
    ->output();
```

## 介于约束

当 `B2` 的值介于 `60` 和 `90`（含两端）之间时高亮：

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

## 文本包含约束

当 `A2` 中的文本包含 `error` 时高亮：

```php
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_TEXT)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_TEXT_CONTAINING)
   ->valueString('error')
   ->format($highlight);

$fileObject->conditionalFormatCell('A2', $cf);
```
