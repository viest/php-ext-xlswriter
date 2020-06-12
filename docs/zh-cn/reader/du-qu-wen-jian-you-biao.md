# 读取文件（游标）

* 读取文件已支持 `windows` 系统，版本号大于等于 `1.3.4.1`；
* 扩展版本大于等于 `1.2.7`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 编译

编译时需添加 `--enable-reader`

```bash
./configure --enable-reader
```

## 示例一

```php
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

## 示例二

```php
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel($config);

// 导出测试文件
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost'])
    ->output();

// 读取测试文件
$excel->openFile('tutorial.xlsx')
    ->openSheet();

// 此处判断请使用【非全等】运算符进行判断，如果出现空行，则有可能返回空数组
while (($row = $excel->nextRow()) !== NULL) {
    var_dump($row);
}
```

