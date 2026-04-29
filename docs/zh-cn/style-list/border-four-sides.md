# 四边边框

为单元格的上、右、下、左四条边分别指定不同的边框样式。

## **函数原型**

```php
borderOfTheFourSides(
    int $top    = Format::BORDER_NONE,
    int $right  = Format::BORDER_NONE,
    int $bottom = Format::BORDER_NONE,
    int $left   = Format::BORDER_NONE
): self
```

### **int $top / $right / $bottom / $left**

> 四条边的边框样式，使用 `Format::BORDER_*` 常量（例如 `BORDER_THIN`、`BORDER_MEDIUM`、`BORDER_DASHED`）。
> 任意参数省略或传 `null` 时，该方向不绘制边框。

需要为四条边分别指定颜色时，请配合 `borderColorOfTheFourSides()` 使用。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$borderStyle = (new \Vtiful\Kernel\Format($fileHandle))
    ->borderOfTheFourSides(
        \Vtiful\Kernel\Format::BORDER_THIN,    // 上
        \Vtiful\Kernel\Format::BORDER_MEDIUM,  // 右
        \Vtiful\Kernel\Format::BORDER_THIN,    // 下
        \Vtiful\Kernel\Format::BORDER_DASHED   // 左
    )
    ->toResource();

$fileObject->header(['name', 'score'])
    ->setRow('A1', 20, $borderStyle)
    ->output();
```
