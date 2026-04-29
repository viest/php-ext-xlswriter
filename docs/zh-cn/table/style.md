# 样式

Excel 表的样式由两部分组合而成：「类型」（默认 / 浅 / 中 / 深）以及类型内部的「编号」。`Table::style(int $type, int $number)` 接受这两个参数。`Table::name(string $name)` 用于给表起一个唯一的名字，未设置时由 libxlsxwriter 生成 `Table1`、`Table2` ...

## 函数原型

```
\Vtiful\Kernel\Table::style(int $type, int $number): self
\Vtiful\Kernel\Table::name(string $name): self
```

### **int $type**

> 样式类型，取以下常量之一：
>
> * `Table::STYLE_TYPE_DEFAULT`
> * `Table::STYLE_TYPE_LIGHT`
> * `Table::STYLE_TYPE_MEDIUM`
> * `Table::STYLE_TYPE_DARK`

### **int $number**

> 类型内部的编号。Excel 默认提供 1~21 号 Light、1~28 号 Medium、1~11 号 Dark 等多套样式，常用值如 `Light 11`、`Medium 9`。`0` 表示不应用任何样式。

### **string $name**

> 表名称。须以字母或下划线开头，并且在工作簿内唯一；可在公式中通过 `Table1[@Column1]` 引用。

## 示例

```php
$config = ['path' => './'];
$excel  = new \Vtiful\Kernel\Excel($config);

$table = new \Vtiful\Kernel\Table();
$table->name('Performance')
      ->style(\Vtiful\Kernel\Table::STYLE_TYPE_MEDIUM, 9)
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
