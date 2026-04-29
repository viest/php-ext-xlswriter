# Insert comment

## **Function Prototype**

```php
insertComment(int $row, int $column, string $text): self

insertCommentOpt(int $row, int $column, string $text, array $options): self

showComment(): self
```

### **int $row**

> cell row

### **int $column**

> cell column

### **string $text**

> comment text

### **array $options**

> Optional extended comment settings. All keys are optional:
>
> - `author` _string_ — comment author
> - `font_name` _string_ — font name
> - `font_size` _double_ — font size
> - `color` _int_ — background color (`0xRRGGBB` or a `Format::COLOR_*` constant)
> - `x_offset` _int_ — horizontal pixel offset
> - `y_offset` _int_ — vertical pixel offset
> - `x_scale` _double_ — horizontal scale factor
> - `y_scale` _double_ — vertical scale factor
> - `width` _double_ — box width in pixels
> - `height` _double_ — box height in pixels
> - `visible` _int_ — visible by default, 1 = shown / 0 = hidden
> - `start_row` _int_ — anchor row of the box
> - `start_col` _int_ — anchor column of the box

`showComment()` toggles the workbook-wide "show all comments" flag. Call it once to make every comment visible by default.

## Example

```php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$file = $excel->fileName('tutorial.xlsx')
    ->header(['name', 'score']);

$file->insertText(1, 0, 'viest')
     ->insertText(1, 1, 99)
     ->insertComment(1, 1, 'Almost a perfect score!')
     ->insertCommentOpt(1, 0, 'Project owner', [
         'author'    => 'admin',
         'font_name' => 'Arial',
         'font_size' => 10,
         'color'     => \Vtiful\Kernel\Format::COLOR_YELLOW,
         'width'     => 200,
         'height'    => 80,
         'visible'   => 1,
     ])
     ->showComment()
     ->output();
```
