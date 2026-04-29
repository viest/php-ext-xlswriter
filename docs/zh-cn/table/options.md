# 可选项

下列方法用于切换 Excel 表的常用开关。所有方法都接受一个可选的 `bool` 参数，省略时默认开启该选项；显式传 `false` 可以关闭。

## 函数原型

```
\Vtiful\Kernel\Table::noHeaderRow(bool $on = true): self
\Vtiful\Kernel\Table::noAutofilter(bool $on = true): self
\Vtiful\Kernel\Table::noBandedRows(bool $on = true): self
\Vtiful\Kernel\Table::bandedColumns(bool $on = true): self
\Vtiful\Kernel\Table::firstColumn(bool $on = true): self
\Vtiful\Kernel\Table::lastColumn(bool $on = true): self
\Vtiful\Kernel\Table::totalRow(bool $on = true): self
```

### **noHeaderRow()**

> 不显示表头行；调用后，传给 `addTable()` 的区域将被整体视为数据行。

### **noAutofilter()**

> 关闭表头上的自动筛选下拉。

### **noBandedRows()**

> 关闭隔行底纹。

### **bandedColumns()**

> 启用隔列底纹，与隔行底纹互不影响。

### **firstColumn() / lastColumn()**

> 高亮第一列或最后一列，常用于强调维度列或汇总列。

### **totalRow()**

> 在区域最后一行启用汇总行，配合 `columns()` 中的 `total_string` / `total_function` 使用。注意区域必须包含汇总行所在那一行。

## 示例

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$table = new \Vtiful\Kernel\Table();
$table->name('Sales')
      ->totalRow()
      ->bandedColumns()
      ->firstColumn()
      ->columns([
          ['header' => 'Region', 'total_string' => 'Total'],
          ['header' => 'Amount', 'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
      ]);

$excel->fileName('tutorial.xlsx')
    ->data([
        ['East', 100],
        ['West',  90],
        ['',       0], // 留空给汇总行
    ])
    ->addTable('A1:B4', $table)
    ->output();
```
