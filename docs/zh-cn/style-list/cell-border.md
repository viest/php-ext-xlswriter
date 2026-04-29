# 单元格边框

## **函数原型**

```php
border(int $borderStyle): \Vtiful\Kernel\Format
```

## 示例

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$data = [
    ['viest1', 21, 100, "A"],
    ['viest2', 20, 80, "B"],
    ['viest3', 22, 70, "C"],
];

$format = new \Vtiful\Kernel\Format($fileHandle);

// 创建边框样式
$borderStyle = $format
    ->border(\Vtiful\Kernel\Format::BORDER_THIN)
    ->toResource();

$fileObject->header(['name', 'age', 'score', 'level'])
    ->data($data)
    ->setRow('A1', 20, $borderStyle)
    ->output();
```

