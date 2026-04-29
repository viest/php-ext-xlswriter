# Conditional format

Conditional formatting applies styles or visualisations to cells whose values match a rule. Common use cases:

- Highlight cells that match a condition (e.g. paint the background red when the value is greater than 50)
- Draw data bars inside cells to compare magnitudes at a glance
- Apply a colour scale gradient across a range
- Mark buckets with an icon set such as a three-colour traffic light
- Filter cells by built-in rules: text / dates / duplicates / top-N / bottom-N

xlswriter builds a rule with `Vtiful\Kernel\ConditionalFormat`, then attaches it to a worksheet via `Excel::conditionalFormatCell` or `Excel::conditionalFormatRange`.

## Methods

```php
Excel::conditionalFormatCell(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
Excel::conditionalFormatRange(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
```

### **string $rangeA1**

> A single cell or a range in A1 notation, e.g. `"A1"`, `"A1:A10"`, `"B2:D8"`.

### **\Vtiful\Kernel\ConditionalFormat $cf**

> A rule object built with `new \Vtiful\Kernel\ConditionalFormat()` and configured via chained calls. Pass the object itself; there is no `toResource()` step.

## Quick start

The example below highlights every cell in `A2:A4` whose value is greater than `50` with a red background and white font:

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// Style applied when the rule matches
$highlight = (new \Vtiful\Kernel\Format($fileHandle))
    ->background(\Vtiful\Kernel\Format::COLOR_RED)
    ->fontColor(\Vtiful\Kernel\Format::COLOR_WHITE)
    ->toResource();

// Build the rule: cell type + greater than 50 + apply $highlight
$cf = new \Vtiful\Kernel\ConditionalFormat();
$cf->type(\Vtiful\Kernel\ConditionalFormat::TYPE_CELL)
   ->criteria(\Vtiful\Kernel\ConditionalFormat::CRITERIA_GREATER_THAN)
   ->value(50)
   ->format($highlight);

$fileObject->header(['score'])
    ->data([[10], [60], [80]])
    ->conditionalFormatRange('A2:A4', $cf)
    ->output();
```

## Pages in this chapter

| Page | Topic |
|------|-------|
| [Single cell rule](cell.md) | Apply one rule to a single cell |
| [Range rule](range.md) | Apply one rule to an A1 range |
| [Data bar](data-bar.md) | Data bar configuration |
| [Icon set](icons.md) | Icon set configuration |
| [Multi range](multi-range.md) | Apply one rule across non-contiguous ranges |
| [Type constants](type-constants.md) | `TYPE_*` constants |
| [Criteria constants](criteria-constants.md) | `CRITERIA_*` constants |
