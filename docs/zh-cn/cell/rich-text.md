# 插入富文本

在同一个单元格里混排不同样式的文本片段，需要使用 `\Vtiful\Kernel\RichString` 把每一段文字与对应的 `Format` 绑定，再把它们组成数组传给 `insertRichText`。

## **函数原型**

```php
\Vtiful\Kernel\RichString::__construct(string $text, ?resource $formatHandle = null)

insertRichText(int $row, int $column, array $runs, ?resource $formatHandle = null): self
```

### **string $text**

> 单段文字内容

### **resource $formatHandle**

> 该段文字使用的样式句柄，由 `Format::toResource()` 生成。传 `null` 则继承单元格样式。

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **array $runs**

> 由若干 `\Vtiful\Kernel\RichString` 实例组成的数组，按顺序依次拼接。
> 数组中如果出现非 `RichString` 实例的元素会抛出异常。

### **resource $formatHandle**（可选）

> 整个单元格层面的样式（例如对齐、背景色），不传则使用默认样式。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$file       = $excel->fileName('tutorial.xlsx');
$fileHandle = $file->getHandle();

$boldStyle = (new \Vtiful\Kernel\Format($fileHandle))
    ->bold()
    ->toResource();

$redStyle = (new \Vtiful\Kernel\Format($fileHandle))
    ->fontColor(\Vtiful\Kernel\Format::COLOR_RED)
    ->toResource();

$italicStyle = (new \Vtiful\Kernel\Format($fileHandle))
    ->italic()
    ->toResource();

$file->insertRichText(0, 0, [
    new \Vtiful\Kernel\RichString('Hello ',  $boldStyle),
    new \Vtiful\Kernel\RichString('World',   $redStyle),
    new \Vtiful\Kernel\RichString(' from ',  null),
    new \Vtiful\Kernel\RichString('xlswriter', $italicStyle),
])->output();
```
