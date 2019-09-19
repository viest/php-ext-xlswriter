# 忽略空白行

## 测试数据准备

```php
$config = ['path' => './tests'];
$excel  = new \Vtiful\Kernel\Excel($config);

// 写入测试数据
$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['', 'Cost'])
    ->data([
        [],
        ['viest', '']
    ])
    ->output();
```

## 示例一

```php
// 读取全量数据
// 使用 \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW 忽略空白行
$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW)
    ->getSheetData();
```

## 示例二

```php
// 游标模式
// 使用 \Vtiful\Kernel\Excel::SKIP_EMPTY_CELLS 忽略空白行

$data = $excel->openFile('tutorial.xlsx')
    ->openSheet('Sheet1', \Vtiful\Kernel\Excel::SKIP_EMPTY_ROW);

while ($data = $excel->nextRow()) {
    var_dump($data);
}
```