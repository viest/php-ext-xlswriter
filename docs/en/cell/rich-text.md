# Insert rich text

To mix multiple styles inside a single cell, build each fragment with `\Vtiful\Kernel\RichString` (which pairs a piece of text with a `Format` resource) and pass an array of those instances to `insertRichText`.

## **Function Prototype**

```php
\Vtiful\Kernel\RichString::__construct(string $text, ?resource $formatHandle = null)

insertRichText(int $row, int $column, array $runs, ?resource $formatHandle = null): self
```

### **string $text**

> Text of one fragment.

### **resource $formatHandle**

> Style handle returned by `Format::toResource()` for that fragment. Pass `null` to inherit the cell style.

### **int $row**

> cell row

### **int $column**

> cell column

### **array $runs**

> An array of `\Vtiful\Kernel\RichString` instances, concatenated in order.
> Any element that is not a `RichString` instance will trigger an exception.

### **resource $formatHandle** (optional)

> Cell-level format (alignment, background, etc.) applied to the whole cell. Defaults to no style.

## Example

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
