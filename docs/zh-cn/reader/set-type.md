# 全局读取类型

* 读取文件已支持 `windows` 系统，版本号大于等于 `1.3.4.1`；
* 扩展版本大于等于 `1.2.7`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 函数原型

```php
setType(array $type)
```

## 类型数组说明

文档第三列是时间，你需要这样设置类型：

```php
[
    2 => \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
]
```

数组下标 `0` 对应文件 `第一列`；

## 测试数据准备

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Name', 'Age', 'Date'])
    ->data([
        ['Viest', 24]
    ])
    ->insertDate(1, 2, 1568877706)
    ->output();
```

## 示例

```php
$data = $excel->openFile('tutorial.xlsx')
    ->openSheet()
    ->setType([
        \Vtiful\Kernel\Excel::TYPE_STRING,
        \Vtiful\Kernel\Excel::TYPE_INT,
        \Vtiful\Kernel\Excel::TYPE_TIMESTAMP,
    ])
    ->getSheetData();

var_dump($data);
```

## 示例输出

```php
array(2) {
  [0]=>
  array(3) {
    [0]=>
    string(4) "Name"
    [1]=>
    string(3) "Age"
    [2]=>
    string(4) "Date"
  }
  [1]=>
  array(3) {
    [0]=>
    string(5) "Viest"
    [1]=>
    int(24)
    [2]=>
    int(1568877706)
  }
}
```