# 缩进

为单元格内文字添加左侧缩进。

## **函数原型**

```php
indent(int $level): self
```

### **int $level**

> 缩进级别，取值范围 `0..15`，每一级约等于 3 个空格的宽度。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$level1 = (new \Vtiful\Kernel\Format($fileHandle))
    ->indent(1)
    ->toResource();

$level3 = (new \Vtiful\Kernel\Format($fileHandle))
    ->indent(3)
    ->toResource();

$fileObject->header(['title'])
    ->insertText(1, 0, '一级缩进', null, $level1)
    ->insertText(2, 0, '三级缩进', null, $level3)
    ->output();
```
