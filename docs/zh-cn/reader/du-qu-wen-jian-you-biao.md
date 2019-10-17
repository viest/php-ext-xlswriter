# 读取文件（游标）

* 读取文件暂不支持 `windows` 系统；
* 扩展版本大于等于 `1.2.7`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 编译

编译时需添加 `--enable-reader`

```bash
./configure --enable-reader
```

## 示例

```bash
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel($config);

// 导出测试文件
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

// 读取测试文件
$excel->openFile('tutorial.xlsx')
    ->openSheet();

var_dump($excel->nextRow()); // ['Item', 'Cost']
var_dump($excel->nextRow()); // NULL
```

