# 文字颜色

### **函数原型**

```php
fontColor(int $color)
```

#### **int $color**

> RGB 十六进制值

### 示例

```php
$config     = ['path' => './tests'];
$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject = $fileObject->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$format     = new \Vtiful\Kernel\Format($fileHandle);
$colorStyle = $format->fontColor(0xFF0000)->toResource();
// 或 
// $colorStyle = $format->fontColor(\Vtiful\Kernel\Format::COLOR_ORANGE)->toResource();

$filePath = $fileObject->header(['name', 'age'])
    ->data([
        ['viest', 21],
        ['wjx',   21]
    ])
    ->setRow('A1', 50, $colorStyle)
    ->output();

var_dump($filePath);
```

