# 固定内存模式

## **内存**

最大内存使用量 = 最大一行的数据占用量

## **注意**

固定内存模式下，单元格按行落盘，如果当前操作的行已落盘则无法进行任何修改；

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

$fileObject->setRow('A1', 10, $boldStyle) // 写入数据前设置行样式
    ->header(['name', 'age'])
    ->data([['viest', 21]])
    ->output();
```

