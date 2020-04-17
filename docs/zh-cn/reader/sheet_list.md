# 工作表列表

* 读取文件已支持 `windows` 系统，版本号大于等于 `1.3.4.1`；
* 扩展版本大于等于 `1.3.2`；
* PECL 安装时将会提示是否开启读取功能，请键入 `yes`；

## 函数原型

```php
sheetList(): array
```

## 示例

```php
$excel = new \Vtiful\Kernel\Excel(['path' => './tests']);

// 构建示例文件
$filePath = $excel
    // 第一个工作表
    ->fileName('tutorial.xlsx', 'test1')
    ->header(['sheet'])
    ->data([['test1']])

    // 第二个工作表
    ->addSheet('test2')
    ->header(['sheet'])
    ->data([['test2']])

    ->output();

// 打开示例文件
$sheetList = $excel->openFile('tutorial.xlsx')
    ->sheetList();

foreach ($sheetList as $sheetName) {
    echo 'Sheet Name:' . $sheetName . PHP_EOL;

    // 通过工作表名称获取工作表数据
    $sheetData = $excel
        ->openSheet($sheetName)
        ->getSheetData();

    var_dump($sheetData);
}
```

## 示例输出

```php
Sheet Name:test1
array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(5) "sheet"
  }
  [1]=>
  array(1) {
    [0]=>
    string(5) "test1"
  }
}
Sheet Name:test2
array(2) {
  [0]=>
  array(1) {
    [0]=>
    string(5) "sheet"
  }
  [1]=>
  array(1) {
    [0]=>
    string(5) "test2"
  }
}
```
