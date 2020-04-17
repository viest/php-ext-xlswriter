# 按数据类型读取文件

* 读取文件已支持 `windows` 系统，版本号大于等于 `1.3.4.1`；
* 扩展版本大于等于 `1.2.7`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 编译

编译时需添加 `--enable-reader`

```bash
./configure --enable-reader
```

## 类型数组说明

文档第三列是时间，你需要这样设置类型：

```php
[
    2 => \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
]
```

数组下标 `0` 对应文件 `第一列`；

## 示例

```bash
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel($config);

// 导出测试文件
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

// 打开测试文件 和 第一个工作表
$excel->openFile('tutorial.xlsx')
    ->openSheet();

// 读取每行单元格数据时，可指定每个单元格数据类型进行读取
var_dump($excel->nextRow([
    \Vtiful\Kernel\Excel::TYPE_STRING,
    \Vtiful\Kernel\Excel::TYPE_STRING
]));
```

