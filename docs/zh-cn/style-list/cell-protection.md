# 单元格保护

控制单元格的锁定状态以及公式的可见性。这三个标志只有在工作表通过 `protection()` 启用保护后才会生效——单纯调用 `locked()` / `unlocked()` 不会让 Excel 进入保护模式。

## **函数原型**

```php
locked(): self

unlocked(): self

hidden(): self
```

### **locked()**

> 将单元格标记为"锁定"。这是 Excel 的默认行为，需要明确想覆盖 `unlocked()` 的效果时调用。

### **unlocked()**

> 将单元格标记为"未锁定"，即使工作表启用保护后这些单元格仍然可以被编辑。

### **hidden()**

> 在工作表启用保护后，隐藏单元格里的公式（公式栏不会再显示原文）。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$editable = (new \Vtiful\Kernel\Format($fileHandle))
    ->unlocked()
    ->toResource();

$secret = (new \Vtiful\Kernel\Format($fileHandle))
    ->locked()
    ->hidden()
    ->toResource();

$fileObject->header(['input', 'formula'])
    ->insertText(1, 0, '可编辑',  null, $editable)
    ->insertFormula(1, 1, '=A2*2', $secret)
    ->protection()
    ->output();
```
