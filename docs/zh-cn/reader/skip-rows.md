# 跳过指定行

* 读取文件已支持 `windows` 系统，版本号大于等于 `1.3.4.1`；
* 扩展版本大于等于 `1.2.7`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 测试数据准备

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

// 写入测试数据
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['', 'Cost'])
    ->data([
        [],
        ['viest', '']
    ])
    ->output();
```

## 示例一

```php
// 读取全量数据, 并忽略第一行
$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1')
    ->setSkipRows(1)
    ->getSheetData();
```

## 示例二

```php
// 读取全量数据, 并忽略第一行 和 第二行
$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1')
    ->setSkipRows(2)
    ->getSheetData();
```