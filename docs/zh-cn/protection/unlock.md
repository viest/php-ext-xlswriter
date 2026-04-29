# 解除编辑保护

当工作表具有编辑保护时，可以使用本样式将特定行、列、单一单元格解除编辑保护，使其可以任意编辑。

## **函数原型**

```php
unlocked();
```

## 示例

```php
$config = [
    'path' => './'
];

$fileObject  = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

// 创建解除保护样式
$format = new \Vtiful\Kernel\Format($fileHandle);
$unlockedStyle = $format->unlocked()->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['wjx',   21]
    ])
    ->setRow('A2', 50, $unlockedStyle) // 将工作表中的第二行解除编辑保护
    ->protection() // 工作表添加编辑保护
    ->output();
```

