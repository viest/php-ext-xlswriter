# 图表

通过 `Vtiful\Kernel\Chart` 构造图表对象，调用 `series()` / `categories()` 注入数据后插入工作表。

> 请先把 `Chart` 实例赋值给变量，再调用工作簿的 `output()`。如果直接链式调用 `(new Chart())->series()->toResource()`，会触发重复插入告警。

本节内容：

* [图表类型常量](chart-type-constants.md)
* [填充数据](data-input.md)
* 示例：[圆环图](doughnut.md)、[面积图](area.md)、[直方图](histogram.md)、[条形图](bar.md)
