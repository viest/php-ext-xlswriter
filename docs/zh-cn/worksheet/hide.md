# 隐藏当前工作表

## **函数原型**

```php
setCurrentSheetHide(): self
```

## **实例**

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

$excel->fileName('hide.xlsx', 'sheet1') // 初始化文件，同时初始化第一张工作表 sheet1
    ->header(['sheet1'])    // 在 sheet1 工作表中插入数据
    ->addSheet('sheet2')    // 添加新工作表 sheet2，并设置当前活动工作表为 sheet2
    ->setCurrentSheetHide() // 当前活动工作表为 sheet2，隐藏 sheet2 工作表
    ->output();
```