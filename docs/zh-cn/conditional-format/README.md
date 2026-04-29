# 条件格式

条件格式（Conditional Format）允许根据单元格的值自动应用不同的样式或可视化效果。常见的使用场景包括：

- 突出显示满足某条件的单元格（例如：`数值大于 50` 时填充红色背景）
- 在数值区域内绘制数据条（Data Bar）以直观对比大小
- 使用色阶（Color Scale）渐变显示数据分布
- 使用图标集（Icon Set）标记区间，例如三色交通灯
- 通过文本/日期/重复值/排名等内置规则筛选目标单元格

xlswriter 通过 `Vtiful\Kernel\ConditionalFormat` 类构造一条规则，再调用 `Excel::conditionalFormatCell` 或 `Excel::conditionalFormatRange` 把规则应用到工作表上。

## 函数原型

```php
Excel::conditionalFormatCell(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
Excel::conditionalFormatRange(string $rangeA1, \Vtiful\Kernel\ConditionalFormat $cf): self
```

### **string $rangeA1**

> 单元格地址或区域，使用 A1 表示法。例如 `"A1"`、`"A1:A10"`、`"B2:D8"`。

### **\Vtiful\Kernel\ConditionalFormat $cf**

> 通过 `new \Vtiful\Kernel\ConditionalFormat()` 创建并链式配置好的规则对象。直接传入对象本身，无需调用 `toResource()`。

## 快速开始

下面的例子在 `A2:A4` 三个单元格上应用一条「数值大于 50 时填充红底白字」的规则：

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// 满足条件时使用的样式
$highlight = (new \Vtiful\Kernel\Format($fileHandle))
    ->background(\Vtiful\Kernel\Format::COLOR_RED)
    ->fontColor(\Vtiful\Kernel\Format::COLOR_WHITE)
    ->toResource();

// 构造规则：单元格类型 + 大于 50 + 应用 $highlight
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

## 章节列表

| 页面 | 内容 |
|-----|------|
| [单元格规则](cell.md) | 在单个单元格上应用条件规则 |
| [区域规则](range.md) | 在一个 A1 区域上应用条件规则 |
| [数据条](data-bar.md) | Data Bar 数据条相关方法 |
| [图标集](icons.md) | Icon Set 图标集相关方法 |
| [多区域](multi-range.md) | 一条规则同时应用到多个不连续区域 |
| [类型常量](type-constants.md) | `TYPE_*` 常量列表 |
| [条件常量](criteria-constants.md) | `CRITERIA_*` 常量列表 |
