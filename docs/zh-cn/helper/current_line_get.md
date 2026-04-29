# 获取当前写入行号

```php
$config = [
    'path' => './tests'
];

$fileObject = new \Vtiful\Kernel\Excel($config);

$fileObject->fileName('tutorial.xlsx')
    ->header(['name', 'age'])
    ->data([
        ['viest', 21],
    ]);

var_dump($fileObject->getCurrentLine()); // 2
```