# 网格线

## **函数原型**

```php
gridline(int $option = \Vtiful\Kernel\Excel::GRIDLINES_SHOW_ALL): self
```

## 网格线类型

```php
const GRIDLINES_HIDE_ALL    = 0; // 隐藏 屏幕网格线 和 打印网格线
const GRIDLINES_SHOW_SCREEN = 1; // 显示屏幕网格线
const GRIDLINES_SHOW_PRINT  = 2; // 显示打印网格线
const GRIDLINES_SHOW_ALL    = 3; // 显示 屏幕网格线 和 打印网格线
```

## **实例**

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$fileObject = $excel->fileName("tutorial01.xlsx");

$fileObject->header(['name', 'age'])
    ->gridline(\Vtiful\Kernel\Excel::GRIDLINES_HIDE_ALL) // 设置工作表网格线
    ->data([
        ['viest', 21],
        ['viest', 22],
        ['viest', 23],
    ])
    ->output();
```

