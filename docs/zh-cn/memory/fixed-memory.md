# 固定内存模式

## **内存**

最大内存使用量 = 最大一行的数据占用量

## **注意**

固定内存模式下，单元格按行落盘，如果当前操作的行已落盘则无法进行任何修改；

## **函数原型**

```php
constMemory(string $fileName[, string $sheetName = 'Sheet1', bool $enableZip64 = true]);
```

## 提示

WPS 需要关闭 zip64，否则打开文件可能报文件损坏；

该问题已反馈给WPS，修复进度未知，并且该问题在出现在iOS、Mac、Windows等平台，Android平台无异常。

## 示例

### 开启 ZIP64

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

// ConstMemory 默认开启 ZIP64
$fileObject = $excel->constMemory('tutorial01.xlsx');
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->setRow('A1', 10, $boldStyle) // 写入数据前设置行样式
    ->header(['name', 'age'])
    ->data([['viest', 21]])
    ->output();
```

### 关闭 ZIP64

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

// 第三个参数 False 即为关闭 ZIP64
$fileObject = $excel->constMemory('tutorial01.xlsx', NULL, false);
$fileHandle = $fileObject->getHandle();

$format    = new \Vtiful\Kernel\Format($fileHandle);
$boldStyle = $format->bold()->toResource();

$fileObject->setRow('A1', 10, $boldStyle) // 写入数据前设置行样式
    ->header(['name', 'age'])
    ->data([['viest', 21]])
    ->output();
```

