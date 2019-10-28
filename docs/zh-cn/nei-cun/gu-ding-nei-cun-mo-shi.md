# 固定内存模式

## **内存**

最大内存使用量 = 最大一行的数据占用量

## **函数原型**

```php
constMemory(string $fileName);
```

## 示例

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->constMemory('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->header(['name', 'age'])
    ->data([['viest', 21]])
    ->setRow('A1', 10, $boldStyle)
    ->output();
```

