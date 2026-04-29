# 检查工作表是否存在

## 函数原型

```php
existSheet(string $sheetName): array
```

## 示例

```php
$config = ['path' => './tests'];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject->fileName('tutorial.xlsx')
    // 添加工作表 twoSheet
    ->addSheet('twoSheet');

var_dump($fileObject->existSheet('twoSheet'));
var_dump($fileObject->existSheet('notFoundSheet'));
```

## 示例输出

```php
bool(true)
bool(false)
```
