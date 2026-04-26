# Chart

Charts are built with `Vtiful\Kernel\Chart`, fed series via `series()` / `categories()`, then inserted into a worksheet.

> Build the `Chart` instance into a variable first, then call `output()` on the workbook. Chaining `(new Chart())->series()->toResource()` inline triggers a double-insert warning.

Pages in this section:

* [Chart type constants](chart-type-constants.md)
* [Data input](data-input.md)
* Examples: [doughnut](doughnut.md), [area](area.md), [histogram](histogram.md), [bar](bar.md)
