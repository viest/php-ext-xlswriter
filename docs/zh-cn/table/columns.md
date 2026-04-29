# 列定义

`Table::columns(array $columns)` 用来配置表中每一列的元数据。`$columns` 是一个数组，每个元素为关联数组，键名描述该列的属性；未列出的列会按 libxlsxwriter 的默认值（`Column1`、`Column2` ...）处理。

## 函数原型

```
\Vtiful\Kernel\Table::columns(array $columns): self
```

### **array $columns**

> 列定义数组，每一项支持以下键：
>
> * **header** *(string)* — 列标题，覆盖默认的 `ColumnN`。
> * **formula** *(string)* — 该列每行使用的公式，例如 `'=SUM(Table1[@[Q1]:[Q4]])'`。
> * **total_string** *(string)* — 汇总行该列显示的字符串，例如 `'Total'`。
> * **total_function** *(int)* — 汇总行该列使用的函数，对应 `Table::FUNCTION_*` 常量。
> * **total_value** *(float)* — 汇总行的数值（一般由 `total_function` 自动算出，如需固定值可设此项）。
> * **format** *(Format|resource)* — 该列单元格的格式对象，可传 `Format` 实例或其资源句柄。
> * **header_format** *(Format|resource)* — 该列表头的格式对象。

`Table::FUNCTION_*` 常量包括：`FUNCTION_NONE`、`FUNCTION_AVERAGE`、`FUNCTION_COUNT_NUMS`、`FUNCTION_COUNT`、`FUNCTION_MAX`、`FUNCTION_MIN`、`FUNCTION_STD_DEV`、`FUNCTION_SUM`、`FUNCTION_VAR`。

## 示例

下面构建一张含 4 个季度数据 + 汇总行的销售表。

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileHandle = $excel->fileName('tutorial.xlsx')->getHandle();
$bold       = new \Vtiful\Kernel\Format($fileHandle);
$boldFmt    = $bold->bold()->toResource();

$table = new \Vtiful\Kernel\Table();
$table->name('Sales')
      ->totalRow()
      ->columns([
          ['header' => 'Region', 'header_format' => $boldFmt, 'total_string' => 'Total'],
          ['header' => 'Q1',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
          ['header' => 'Q2',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
          ['header' => 'Q3',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
          ['header' => 'Q4',     'total_function' => \Vtiful\Kernel\Table::FUNCTION_SUM],
      ]);

$excel->data([
        ['East',  100, 110, 120, 130],
        ['West',   90,  95, 100, 105],
        ['North', 200, 210, 220, 230],
        // 最后一行留给汇总行
        ['',        0,   0,   0,   0],
    ])
    ->addTable('A1:E5', $table)
    ->output();
```
