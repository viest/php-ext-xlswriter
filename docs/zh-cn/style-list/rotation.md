# 文字旋转

旋转单元格内文字的显示角度。

## **函数原型**

```php
rotation(int $angle): self
```

### **int $angle**

> 旋转角度，单位为度。
>
> - 正值表示逆时针旋转，例如 `45` 表示向左上倾斜 45°；
> - 负值表示顺时针旋转，例如 `-45` 表示向右下倾斜 45°；
> - 取值范围 `-90..90`；
> - 特殊值 `270` 表示文字纵向（每个字符堆叠在前一个字符下方）。
>
> 超出有效范围的值会被忽略。

## 示例

```php
$config = [
    'path' => './tests'
];

$excel      = new \Vtiful\Kernel\Excel($config);
$fileObject = $excel->fileName('tutorial.xlsx');
$fileHandle = $fileObject->getHandle();

$rotate45 = (new \Vtiful\Kernel\Format($fileHandle))
    ->rotation(45)
    ->toResource();

$vertical = (new \Vtiful\Kernel\Format($fileHandle))
    ->rotation(270)
    ->toResource();

$fileObject->header(['横排', '斜排', '纵排'])
    ->insertText(1, 0, 'normal')
    ->insertText(1, 1, 'tilted',   null, $rotate45)
    ->insertText(1, 2, 'vertical', null, $vertical)
    ->output();
```
