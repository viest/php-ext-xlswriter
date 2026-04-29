# 读取文件

读取功能可以打开已有 XLSX 文件，并通过两种方式访问数据：全量读取（一次性返回二维数组）或游标读取（逐行返回，内存恒定）。除了单元格内容外，还可以读取样式、合并单元格、超链接、数据验证、条件格式、页面设置、定义名称以及图片、批注、图表等元数据。

读取功能需要在编译扩展时启用 `--enable-reader`（PECL 包默认开启）。

本节内容：

* 迭代：[工作表列表](sheet_list.md)、[工作表列表（含元数据）](sheet-list-with-meta.md)、[全量读取](read-file-full.md)、[游标读取](read-file-cursor.md)、[单元格回调模式](cell_callback.md)
* 过滤：[跳过指定行](skip-rows.md)、[忽略空白单元格](skip-empty-cells.md)、[忽略空白行](skip-empty-row.md)、[跳过常量](skip-const.md)、[设置全局读取类型](set-type.md)、[按数据类型读取](data-type.md)、[数据类型常量](data-type-const.md)
* 单元格：[读取公式](next-row-with-formula.md)、[读取富文本](next-row-rich.md)、[读取样式](get-style-format.md)
* 工作表元数据：[合并单元格](get-merged-cells.md)、[超链接](get-hyperlinks.md)、[工作表保护](get-sheet-protection.md)、[行 / 列设置](row-column-options.md)、[条件格式](get-conditional-formats.md)、[数据验证](get-data-validations.md)、[自动筛选](get-auto-filter.md)、[定义名称](get-defined-names.md)、[页面设置](get-page-setup.md)、[公式 AST](get-formula-ast.md)
* 嵌入对象：[遍历图片](iterate-images.md)、[遍历批注](iterate-comments.md)、[遍历图表](iterate-charts.md)
