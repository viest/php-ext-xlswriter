# 单元格回调模式读取

* 读取文件已支持 `windows` 系统，版本号大于等于 `1.3.4.1`；
* 扩展版本大于等于 `1.2.7`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 优势

`最大内存` == `最大单元格数据体积`

该模式可满足 `xlsx ` 大文件读取

## 函数原型

```php
nextCellCallback(callable $callback, string $sheetName = NULL): void
```

## 回调函数原型

```php
function(int $row, int $cell, string|double|int $data)
```

## 单行结束标识

每行末尾，将会`附加一次回调`，并传递 `XLSX_ROW_END` 为当前行结束标识。

## 示例

```php
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel($config);
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost'])
    ->data([
        ['Item_1', 'Cost_1'],
    ])
    ->output();

$excel->openFile('tutorial.xlsx')->nextCellCallback(function ($row, $cell, $data) {
    echo 'cell:' . $cell . ', row:' . $row . ', value:' . $data . PHP_EOL;
});
```

## 示例输出

```php
cell:0, row:0, value:Item
cell:1, row:0, value:Cost
cell:1, row:0, value:XLSX_ROW_END  // 结束标识
cell:0, row:1, value:Item_1
cell:1, row:1, value:Cost_1
cell:1, row:1, value:XLSX_ROW_END // 结束标识
```