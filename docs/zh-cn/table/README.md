# Excel 表

Excel 表（Table）会把工作表中的某一块区域标识为一张「智能表」，自带表头、自动筛选、隔行底纹、汇总行等特性。表本身仍然由普通单元格构成，只是 `xl/tables/tableN.xml` 中多了一份元数据。

整体流程分两步：

1. 通过 `Vtiful\Kernel\Table` 构造表选项；
2. 调用 `Excel::addTable(string $rangeA1, ?Table $opts = null)` 把表应用到指定区域。

> 注意：`Table` 不需要调用 `toResource()`，直接把 `Table` 实例作为第二个参数传给 `addTable()` 即可；省略第二个参数时使用默认表样式。

## 函数原型

```
\Vtiful\Kernel\Excel::addTable(string $rangeA1, ?\Vtiful\Kernel\Table $opts = null): self
```

### **string $rangeA1**

> 表所覆盖的区域，使用 A1 表示法，例如 `A1:D11`。区域至少包含表头与一行数据。

### **?Table $opts**

> 可选的 `Table` 选项对象。为 `null` 时使用 libxlsxwriter 默认配置。

## 子章节

* [列定义](columns.md)
* [样式](style.md)
* [可选项](options.md)

## 示例

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$table = new \Vtiful\Kernel\Table();
$table->name('Performance')
      ->style(\Vtiful\Kernel\Table::STYLE_TYPE_LIGHT, 11)
      ->columns([
          ['header' => 'Name'],
          ['header' => 'Score'],
      ]);

$excel->fileName('tutorial.xlsx')
    ->data([
        ['Alice', 90],
        ['Bob',   80],
    ])
    ->addTable('A1:B3', $table)
    ->output();
```
