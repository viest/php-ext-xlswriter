# 全局读取类型

## 函数原型

```php
setType(array $type)
```

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