# 页面设置

* 扩展版本大于等于 `1.6.0`；
* 编译时需添加 `--enable-reader`；

读取当前工作表的页面设置，包括边距、纸张、缩放、页眉页脚等。

## 函数原型

```php
getPageSetup(): ?array
```

## 返回值

工作表未声明任何页面设置时返回 `null`，否则返回如下分组字段：

**边距**

* `margins` *(array\|null)* — `{left, right, top, bottom, header, footer}`，单位英寸。

**打印 / 缩放**

* `paper_size` *(int\|null)* — 纸张代码；
* `fit_to_width` *(int\|null)* — 适合宽度页数；
* `fit_to_height` *(int\|null)* — 适合高度页数；
* `scale` *(int\|null)* — 缩放百分比；
* `orientation` *(string\|null)* — `portrait` / `landscape`；
* `horizontal_dpi` *(int\|null)* — 横向 DPI；
* `vertical_dpi` *(int\|null)* — 纵向 DPI；
* `first_page_number` *(int\|null)* — 起始页码；
* `use_first_page_number` *(bool)* — 是否使用自定义起始页码。

**居中 / 网格**

* `print_horizontal_centered` *(bool)* — 水平居中；
* `print_vertical_centered` *(bool)* — 垂直居中；
* `print_grid_lines` *(bool)* — 打印网格线；
* `print_headings` *(bool)* — 打印行列标题。

**页眉页脚**

* `odd_header` / `odd_footer` *(string\|null)* — 奇数页页眉 / 页脚；
* `even_header` / `even_footer` *(string\|null)* — 偶数页页眉 / 页脚；
* `first_header` / `first_footer` *(string\|null)* — 首页页眉 / 页脚；
* `different_odd_even` *(bool)* — 奇偶页不同；
* `different_first` *(bool)* — 首页不同；
* `scale_with_doc` *(bool)* — 与文档一同缩放；
* `align_with_margins` *(bool)* — 与页边距对齐。

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$ps = $excel->openFile('source.xlsx')
    ->openSheet()
    ->getPageSetup();

var_export($ps);
```
