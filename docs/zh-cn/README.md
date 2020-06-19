# 中文简体

### 为什么内存优化模式内存占用少？

当开启内存优化模式时，单元格将根据行落地磁盘，内存中只保留最新一行数据，所以内存优化模式最大内存占用等于数据最多一行的内存用量。

源码解析：

```c
if (!self->optimize) {
    // 未开启内存优化模式......
} else {
    if (row_num < self->optimize_row->row_num) {
        // 当请求行号小于内存行号，无法获取行数据，故无法进行折返写入。
        return NULL;
    } else if (row_num == self->optimize_row->row_num) {
        // 获取当前行信息，提供进行单元格写入。
        return self->optimize_row;
    } else {
        // 当请求行号大于内存行号，则将内存中的行数据刷入磁盘。
        // 重置内存中行数据，并赋予最新行号。
        lxw_worksheet_write_single_row(self);
        row = self->optimize_row;
        row->row_num = row_num;
        return row;
    }
}
```

### 为什么我本地导出比评测慢？

写入流程：创建临时文件 -> 写入XML文件 -> 结束写入 -> 分段读取临时文件并写入压缩文件；

从写入流程看，可将 xlsx 文件写入大致分为两大块，XML写入 和 打包压缩文件。虽然写入XML文件这里可以利用操作系统缓冲区，
减少磁盘IO次数，但最后读取临时文件并写入压缩文件时存在将磁盘中的数据分段重新载入内存，并再次写入压缩文件，在这个过程中对主机的硬盘读取性能以及CPU
处理压缩算法性能密切相关。