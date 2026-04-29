# 快速上手

通过两个最小示例覆盖最常见的使用场景：

* [导出文件](create.md) —— 从零开始创建一个 XLSX 文件。
* [读取文件](reader.md) —— 打开已有 XLSX 并按行迭代数据。
* [手动销毁资源](close.md) —— 在长生命周期脚本中显式释放底层资源。

示例均使用 `Vtiful\Kernel\Excel` 类。读取相关 API 需要在编译扩展时启用 `--enable-reader`（PECL 包默认开启）。
