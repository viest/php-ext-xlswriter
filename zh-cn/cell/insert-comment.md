# 插入批注

## **函数原型**

```php
insertComment(int $row, int $column, string $text): self

insertCommentOpt(int $row, int $column, string $text, array $options): self

showComment(): self
```

### **int $row**

> 单元格所在行

### **int $column**

> 单元格所在列

### **string $text**

> 批注内容

### **array $options**

> 批注扩展选项，所有键均为可选：
>
> - `author` _string_ —— 批注作者
> - `font_name` _string_ —— 批注字体
> - `font_size` _double_ —— 批注字号
> - `color` _int_ —— 批注背景色（`0xRRGGBB` 或 `Format::COLOR_*`）
> - `x_offset` _int_ —— X 方向像素偏移
> - `y_offset` _int_ —— Y 方向像素偏移
> - `x_scale` _double_ —— 横向缩放比例
> - `y_scale` _double_ —— 纵向缩放比例
> - `width` _double_ —— 批注框宽度（像素）
> - `height` _double_ —— 批注框高度（像素）
> - `visible` _int_ —— 是否默认显示，1 显示 / 0 隐藏
> - `start_row` _int_ —— 批注框起始行
> - `start_col` _int_ —— 批注框起始列

`showComment()` 用于打开整个工作簿的批注默认显示开关，调用一次即可对所有批注生效。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel = new \Vtiful\Kernel\Excel($config);

$file = $excel->fileName('tutorial.xlsx')
    ->header(['name', 'score']);

$file->insertText(1, 0, 'viest')
     ->insertText(1, 1, 99)
     ->insertComment(1, 1, '满分接近，继续努力')
     ->insertCommentOpt(1, 0, '主要负责人', [
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
